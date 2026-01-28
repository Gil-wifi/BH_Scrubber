#!/usr/bin/env python3
"""
BH_Scrubber - Bank Holiday Data Aggregation Tool

Scrapes public holiday data from officeholidays.com and populates
a generated ODS spreadsheet with holiday names and dates.

Features:
- Generates Solar or Fiscal calendars based on user input.
- Scrapes multi-year data if fiscal year overlaps.
- Applies Green (Supported) / Grey (Unsupported) row styling.
- Distinguishes National (Amber) vs Regional (Pink) holidays.

Version: 2.1
Author: Gil Alowaifi
Date: 2026-01-28
"""

__version__ = "2.1"

import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
from datetime import datetime, timedelta
import urllib.request
from html.parser import HTMLParser
import time
import re

class OfficeHolidaysParser(HTMLParser):
    """Parses officeholidays.com HTML to extract holiday names, dates, and types."""
    
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_row = False
        self.current_cell_idx = 0
        self.holidays = []  # List of (date_obj, name, is_national) - Note: date_obj now, not str
        
        self.temp_date = None
        self.temp_name = None
        self.temp_type = None
        
        self.recording_name = False
        self.recording_type = False
        self.current_name_text = ""
        self.current_type_text = ""

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        class_name = attrs_dict.get('class', '')
        
        if tag == 'table' and 'country-table' in class_name:
            self.in_table = True
            
        if self.in_table and tag == 'tr':
            self.in_row = True
            self.current_cell_idx = 0
            self.temp_date = None
            self.temp_name = None
            self.temp_type = None
            self.current_type_text = ""
            
        if self.in_row and tag == 'td':
            self.current_cell_idx += 1
            # Cell 4 is the holiday type column
            if self.current_cell_idx == 4:
                self.recording_type = True
                self.current_type_text = ""
            
        if self.in_row and tag == 'time':
            # Date usually in datetime attr YYYY-MM-DD
            self.temp_date = attrs_dict.get('datetime')
            
        if self.in_row and tag == 'a' and 'country-listing' in class_name:
            self.recording_name = True
            self.current_name_text = ""

    def handle_endtag(self, tag):
        if tag == 'table':
            self.in_table = False
        if tag == 'td':
            if self.recording_type:
                self.recording_type = False
                self.temp_type = self.current_type_text.strip()
        if tag == 'tr':
            self.in_row = False
            if self.temp_date and self.current_name_text:
                try:
                    dt = datetime.strptime(self.temp_date, "%Y-%m-%d")
                    # Determine if regional/local
                    type_lower = (self.temp_type or '').lower()
                    
                    is_regional = 'regional' in type_lower or 'local' in type_lower
                    is_national = not is_regional
                    
                    self.holidays.append((dt, self.current_name_text.strip(), is_national))
                except ValueError:
                    pass  # Invalid date format
        if tag == 'a':
            self.recording_name = False

    def handle_data(self, data):
        if self.recording_name:
            self.current_name_text += data
        if self.recording_type:
            self.current_type_text += data


class HolidayScraper:
    def fetch_holidays(self, base_url, years):
        """Fetches holidays for one or more years."""
        all_holidays = []
        
        # Clean URL: ensure we don't double up year logic
        # Typically .../country/sub/year -> strip year
        # Heuristic: if url ends with /YYYY, strip it
        clean_url = base_url.rstrip('/')
        if re.search(r'/\d{4}$', clean_url):
            clean_url = clean_url.rsplit('/', 1)[0]
            
        seen_urls = set()
        
        for year in years:
            url = f"{clean_url}/{year}"
            if url in seen_urls: continue
            seen_urls.add(url)
            
            # print(f"    Fetching {url}...") 
            try:
                req = urllib.request.Request(
                    url, 
                    data=None, 
                    headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                    }
                )
                with urllib.request.urlopen(req, timeout=10) as response:
                    html_content = response.read().decode('utf-8')
                    
                parser = OfficeHolidaysParser()
                parser.feed(html_content)
                all_holidays.extend(parser.holidays)
                
            except urllib.error.HTTPError as e:
                if e.code == 404 and year > datetime.now().year:
                    print(f"    Note: Data for {year} not yet published.")
                else:
                    print(f"    Scrape Warning ({url}): {e}")
            except Exception as e:
                print(f"    Scrape Error ({url}): {e}")
                
        return all_holidays

class ODSHandler:
    def __init__(self, filename):
        self.filename = filename
        self.ns = {
            'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
            'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
            'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
            'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
            'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
            'xlink': 'http://www.w3.org/1999/xlink'
        }
        for prefix, uri in self.ns.items():
            ET.register_namespace(prefix, uri)
            
        self.content_tree = None
        self.root = None
        self.spreadsheet = None
        self.table = None
        self.date_map = {} # date_obj -> col_idx
        
        self.load()

    def load(self):
        with zipfile.ZipFile(self.filename, 'r') as z:
            with z.open("content.xml") as f:
                self.content_tree = ET.parse(f)
                self.root = self.content_tree.getroot()

        self.spreadsheet = self.root.find('.//office:body/office:spreadsheet', self.ns)
        self.table = self.spreadsheet.find('table:table', self.ns)

    def generate_calendar_headers(self, start_date, duration_days):
        """
        Extends the spreadsheet to include dates from start_date for duration_days.
        Populates Row 1 (Header) with 'Thu 01/01/26' format dates.
        """
        rows = self.table.findall('table:table-row', self.ns)
        if not rows: return
        
        header_row = rows[0]
        # Current length. We assume cols 0-3 (4 data columns) exist.
        # We append from column 4.
        
        current_date = start_date
        
        # We need to know where we start.
        # Check current cells in header
        existing_cells = self.expand_row(header_row)
        start_col_idx = len(existing_cells)
        
        # We assume clean slate for calendar part or appending?
        # Specification implies we are building on the template that has just countries.
        # So we append.
        
        print(f"Generating calendar from {start_date.strftime('%Y-%m-%d')} for {duration_days} days...")
        
        for i in range(duration_days):
            # Format: Thu 01/01/26
            date_str = current_date.strftime('%a %d/%m/%y')
            
            # Create cell
            cell = ET.Element(f'{{{self.ns["table"]}}}table-cell')
            cell.attrib[f'{{{self.ns["office"]}}}value-type'] = "string"
            p = ET.SubElement(cell, f'{{{self.ns["text"]}}}p')
            p.text = date_str
            
            # Add to header
            header_row.append(cell)
            
            # Map date (stripping time for safety)
            # We map the datetime OBJECT (normalized) to column index
            normalized_date = datetime(current_date.year, current_date.month, current_date.day)
            self.date_map[normalized_date] = start_col_idx + i
            
            current_date += timedelta(days=1)
            
        print(f"Calendar columns added. Total columns: {start_col_idx + duration_days}")

    def expand_row(self, row):
        """Expands a row with repeated columns into a flat list of text values."""
        cells = row.findall('table:table-cell', self.ns)
        row_data = []
        for cell in cells:
            repeated = cell.attrib.get(f'{{{self.ns["table"]}}}number-columns-repeated')
            count = int(repeated) if repeated else 1
            text_node = cell.find('text:p', self.ns)
            text = text_node.text if text_node is not None else ""
            for _ in range(count):
                row_data.append(text)
        return row_data

    def get_countries(self):
        """Returns list of (row_index, country_name). Starting from Row 2 (index 1)."""
        countries = []
        rows = self.table.findall('table:table-row', self.ns)
        for i, row in enumerate(rows):
            if i < 1: continue
            
            # Column D (index 3)
            cells = row.findall('table:table-cell', self.ns)
            col_idx = 0
            country_name = None
            COUNTRY_COL = 3
            
            for cell in cells:
                repeated = cell.attrib.get(f'{{{self.ns["table"]}}}number-columns-repeated')
                count = int(repeated) if repeated else 1
                
                if col_idx <= COUNTRY_COL < col_idx + count:
                    text_node = cell.find('text:p', self.ns)
                    if text_node is not None:
                        link = text_node.find('text:a', self.ns)
                        if link is not None and link.text:
                            country_name = link.text
                        elif text_node.text:
                            country_name = text_node.text
                    break
                col_idx += count
            
            if country_name:
                countries.append((i, country_name))
        return countries

    def get_cell_node(self, row_idx, col_idx, auto_extend=True):
        """Finds or creates cell. Auto-extends row if col_idx is beyond bounds."""
        rows = self.table.findall('table:table-row', self.ns)
        if row_idx >= len(rows): return None
        
        row = rows[row_idx]
        cells = row.findall('table:table-cell', self.ns)
        
        # Calculate current width
        current_width = 0
        for cell in cells:
            repeated = cell.attrib.get(f'{{{self.ns["table"]}}}number-columns-repeated')
            count = int(repeated) if repeated else 1
            current_width += count
            
        # Extend if needed
        if col_idx >= current_width:
            if not auto_extend: return None
            # Append empty cells
            missing = (col_idx - current_width) + 1
            # We can enable 'repeated' for efficiency if missing is large, 
            # but for simplicity let's append one by one or a block?
            # A block is better.
            new_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell')
            if missing > 1:
                new_cell.attrib[f'{{{self.ns["table"]}}}number-columns-repeated'] = str(missing)
            row.append(new_cell)
            
            # Refresh cells list
            cells = row.findall('table:table-cell', self.ns)

        # Proceed to find target
        import copy
        current_col = 0
        for i, cell in enumerate(cells):
            repeated = cell.attrib.get(f'{{{self.ns["table"]}}}number-columns-repeated')
            count = int(repeated) if repeated else 1
            
            if current_col <= col_idx < current_col + count:
                if count > 1:
                    # Split logic
                    pre_count = col_idx - current_col
                    post_count = (current_col + count) - (col_idx + 1)
                    
                    new_cells_list = []
                    base_attribs = dict(cell.attrib)
                    if f'{{{self.ns["table"]}}}number-columns-repeated' in base_attribs:
                        del base_attribs[f'{{{self.ns["table"]}}}number-columns-repeated']
                    
                    if pre_count > 0:
                        pre_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell', base_attribs)
                        pre_cell.attrib[f'{{{self.ns["table"]}}}number-columns-repeated'] = str(pre_count)
                        new_cells_list.append(pre_cell)
                        
                    target_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell', base_attribs)
                    new_cells_list.append(target_cell)
                    
                    if post_count > 0:
                        post_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell', base_attribs)
                        post_cell.attrib[f'{{{self.ns["table"]}}}number-columns-repeated'] = str(post_count)
                        new_cells_list.append(post_cell)

                    row.remove(cell)
                    for nc in reversed(new_cells_list):
                        row.insert(i, nc)
                    return target_cell
                else:
                    return cell
            current_col += count
        return None

    def update_cell_text(self, row_idx, col_idx, text_val, style_name=None):
        cell = self.get_cell_node(row_idx, col_idx, auto_extend=True)
        if cell is not None:
            # Clear text
            for child in list(cell):
                if child.tag == f'{{{self.ns["text"]}}}p':
                    cell.remove(child)
            
            if text_val:
                p = ET.Element(f'{{{self.ns["text"]}}}p')
                p.text = text_val
                cell.append(p)
                cell.attrib[f'{{{self.ns["office"]}}}value-type'] = "string"
            
            if style_name:
                cell.attrib[f'{{{self.ns["table"]}}}style-name'] = style_name

    # --- Style Helpers (Consolidated) ---
    def _create_style(self, name, bg_color, font_size=None):
        auto_styles = self.root.find('office:automatic-styles', self.ns)
        if auto_styles is None:
            auto_styles = ET.SubElement(self.root, f'{{{self.ns["office"]}}}automatic-styles')
            
        for style in auto_styles.findall('style:style', self.ns):
            if style.get(f'{{{self.ns["style"]}}}name') == name:
                return name
                
        style = ET.SubElement(auto_styles, f'{{{self.ns["style"]}}}style', {
            f'{{{self.ns["style"]}}}name': name,
            f'{{{self.ns["style"]}}}family': "table-cell",
            f'{{{self.ns["style"]}}}parent-style-name': "Default"
        })
        ET.SubElement(style, f'{{{self.ns["style"]}}}table-cell-properties', {
            f'{{{self.ns["fo"]}}}background-color': bg_color
        })
        if font_size:
            ET.SubElement(style, f'{{{self.ns["style"]}}}text-properties', {
                f'{{{self.ns["fo"]}}}font-size': font_size
            })
        return name

    def ensure_styles(self):
        self.amber = self._create_style("AmberHoliday", "#ffbf00", "6pt")
        self.pink = self._create_style("PinkHoliday", "#ffb6c1", "6pt")
        self.green = self._create_style("GreenRow", "#ccffcc")
        self.grey = self._create_style("GreyRow", "#808080")

    def apply_row_style(self, row_idx, start_col, end_col, style_name, exclude_styles=None, force=False):
        for c_idx in range(start_col, end_col + 1):
            cell = self.get_cell_node(row_idx, c_idx, auto_extend=True)
            if cell is not None:
                current_style = cell.get(f'{{{self.ns["table"]}}}style-name')
                if not force and exclude_styles and current_style in exclude_styles:
                    continue
                cell.attrib[f'{{{self.ns["table"]}}}style-name'] = style_name

    def save(self):
        # Filename is already set to the output name at init
        temp_zip = self.filename + ".temp_swp"
        
        with zipfile.ZipFile(self.filename, 'r') as zin:
            with zipfile.ZipFile(temp_zip, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == "content.xml":
                        zout.writestr(item, ET.tostring(self.root, encoding='utf-8', xml_declaration=True))
                    else:
                        zout.writestr(item, zin.read(item.filename))
        
        shutil.move(temp_zip, self.filename)
        print(f"Saved to {self.filename}")


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    print("--- Bank Holiday Generator v2.1 ---")
    
    # 1. Inputs
    year_str = input("Enter the target year (e.g., 2026): ").strip()
    year = int(year_str) if year_str else 2026
    
    fiscal_ans = input("Do you require a Fiscal Calendar? (y/n): ").strip().lower()
    is_fiscal = (fiscal_ans == 'y')
    
    offset_weeks = 0
    if is_fiscal:
        off_str = input("Enter the shift in weeks (e.g., 5): ").strip()
        offset_weeks = int(off_str) if off_str else 0

    # 2. Date Calculations
    # Standard: Jan 1 to Dec 31 (365/366 days)
    # Fiscal: Jan 1 + offset -> + Year Duration
    
    base_start = datetime(year, 1, 1)
    
    # Calculate duration of the TARGET year (Is 2026 leap?)
    # Rule: "calendar must span exactly one full year"
    # Logic: If 2026 is 365 days, span 365. If 366, span 366.
    def is_leap(y):
        return (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0)
    
    duration = 366 if is_leap(year) else 365
    
    if is_fiscal:
        start_date = base_start + timedelta(weeks=offset_weeks)
    else:
        start_date = base_start
        
    end_date_inclusive = start_date + timedelta(days=duration - 1)
    print(f"\nCalendar Configuration:")
    print(f"  Start Date: {start_date.strftime('%Y-%m-%d')}")
    print(f"  End Date:   {end_date_inclusive.strftime('%Y-%m-%d')}")
    print(f"  Duration:   {duration} days")
    
    # Define Years present in range (for scraping)
    years_to_scrape = set()
    curr = start_date
    while curr <= end_date_inclusive:
        years_to_scrape.add(curr.year)
        curr = datetime(curr.year, 12, 31) + timedelta(days=1) # Jump to next year roughly
        # Safer: just check start and end year.
    years_to_scrape.add(start_date.year)
    years_to_scrape.add(end_date_inclusive.year)
    print(f"  Scraping Years: {sorted(list(years_to_scrape))}")

    # 3. File Setup
    input_template = "country_List.ods" 
    if not os.path.exists(input_template):
        # Fallback if user hasn't renamed it yet, likely 'BH_List.ods' if 'country_List.ods' missing?
        # But I saw country_List.ods in ls output (as lock file name implied). 
        # Wait, ls output showed 'country_List.ods' exists.
        pass
    
    date_stamp = datetime.now().strftime('%Y%m%d')
    output_filename = f"BH_List_{date_stamp}.ods"
    
    print(f"\nCreating {output_filename} from template...")
    shutil.copyfile(input_template, output_filename)
    
    # 4. Generate & Populate
    ods = ODSHandler(output_filename)
    ods.generate_calendar_headers(start_date, duration)
    ods.ensure_styles()
    
    countries = ods.get_countries()
    scraper = HolidayScraper()
    
    print("\n--- Starting Data Population ---")
    
    cols_count = len(ods.date_map)
    start_col_idx = min(ods.date_map.values()) if ods.date_map else 4
    end_col_idx = max(ods.date_map.values()) if ods.date_map else 4 + duration
    
    status_counts = {'Yes': 0, 'No': 0, 'Blank': 0}

    for row_idx, country_name in countries:
        # Check Status
        status = "Blank"
        c_a = ods.get_cell_node(row_idx, 0, auto_extend=True)
        if c_a is not None:
             p = c_a.find('text:p', ods.ns)
             if p is not None and p.text:
                 t = p.text.strip().lower()
                 if t == 'yes': status = 'Yes'
                 elif t == 'no': status = 'No'
        
        status_counts[status] += 1
        
        # Apply Row Style
        if status == 'Yes':
            ods.apply_row_style(row_idx, 0, end_col_idx, ods.green, exclude_styles=[ods.amber, ods.pink])
        elif status == 'No':
            ods.apply_row_style(row_idx, 0, end_col_idx, ods.grey, force=True)
            continue # Skip scraping
            
        # Get URL
        url_text = ""
        c_d = ods.get_cell_node(row_idx, 3)
        if c_d is not None:
             p = c_d.find('text:p', ods.ns)
             if p is not None:
                 lnk = p.find('text:a', ods.ns)
                 if lnk is not None:
                     url_text = lnk.attrib.get(f'{{{ods.ns["xlink"]}}}href', '')
        
        if url_text and "officeholidays.com" in url_text:
            print(f"Scraping [{country_name}]...")
            holidays = scraper.fetch_holidays(url_text, sorted(list(years_to_scrape)))
            
            count = 0
            for h_date, h_name, is_national in holidays:
                # Check bounds
                if start_date <= h_date <= end_date_inclusive:
                    # Map to col
                    # date_obj to date_obj lookup
                    normalized = datetime(h_date.year, h_date.month, h_date.day)
                    if normalized in ods.date_map:
                        col = ods.date_map[normalized]
                        style = ods.amber if is_national else ods.pink
                        ods.update_cell_text(row_idx, col, h_name, style)
                        count += 1
            print(f"  -> Added {count} holidays.")
            time.sleep(0.2) # Polite delay
            
    ods.save()
    print("\nProcessing Complete!")
