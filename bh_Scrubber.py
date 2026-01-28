#!/usr/bin/env python3
"""
BH_Scrubber - Bank Holiday Data Aggregation Tool

Scrapes public holiday data from officeholidays.com and populates
the BH_List.ods spreadsheet with holiday names and dates.

Version: 1.2
Author: Gil Alowaifi
Date: 2026-01-28
"""

__version__ = "1.2"

import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
from datetime import datetime
import urllib.request
from html.parser import HTMLParser
import time

class OfficeHolidaysParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_table = False
        self.in_row = False
        self.current_cell_idx = 0
        self.holidays = [] # List of (date_str, name)
        
        self.temp_date = None
        self.temp_name = None
        
        self.recording_name = False
        self.current_name_text = ""

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
            
        if self.in_row and tag == 'td':
            self.current_cell_idx += 1
            
        if self.in_row and tag == 'time':
            # Date usually in datetime attr
            self.temp_date = attrs_dict.get('datetime')
            
        if self.in_row and tag == 'a' and 'country-listing' in class_name:
            self.recording_name = True
            self.current_name_text = ""

    def handle_endtag(self, tag):
        if tag == 'table':
            self.in_table = False
        if tag == 'tr':
            self.in_row = False
            if self.temp_date and self.current_name_text:
                # Convert YYYY-MM-DD to DD/MM/YY
                try:
                    dt = datetime.strptime(self.temp_date, "%Y-%m-%d")
                    fmt_date = dt.strftime("%d/%m/%y")
                    self.holidays.append((fmt_date, self.current_name_text.strip()))
                except ValueError:
                    pass # Invalid date format
        if tag == 'a':
            self.recording_name = False

    def handle_data(self, data):
        if self.recording_name:
            self.current_name_text += data

class HolidayScraper:
    def fetch_holidays(self, url):
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
            return parser.holidays
            
        except Exception as e:
            print(f"  Scrape Error ({url}): {e}")
            return []

class ODSHandler:
    def __init__(self, filename="BH_List.ods", year=None):
        self.filename = filename
        self.year = str(year) if year else None
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
        self.styles_tree = None
        
        # Load immediately
        self.load()

    def load(self):
        # We extract content to memory
        with zipfile.ZipFile(self.filename, 'r') as z:
            with z.open("content.xml") as f:
                self.content_tree = ET.parse(f)
                self.root = self.content_tree.getroot()
            try:
                with z.open("styles.xml") as f:
                    self.styles_tree = ET.parse(f)
            except KeyError:
                print("No styles.xml found (might be in content.xml mostly)")

        # Locate Table
        self.spreadsheet = self.root.find('.//office:body/office:spreadsheet', self.ns)
        self.table = self.spreadsheet.find('table:table', self.ns)
        
        # Parse Dates (Header Row 1 - index 0, format: 'Thu 01/01/26')
        self.date_map = {} # "DD/MM/YY" -> Column Index
        rows = self.table.findall('table:table-row', self.ns)
        if len(rows) > 0:
            header_cells = self.expand_row(rows[0]) # Row 1 (index 0)
            for idx, text in enumerate(header_cells):
                # New format: "Thu 01/01/26" - extract date part
                if "/" in text:
                    # Extract just the date portion (after the day abbreviation)
                    parts = text.split()
                    date_part = parts[-1] if len(parts) > 1 else text  # Get "01/01/26"
                    self.date_map[date_part] = idx
                    
                    # Validate Year
                    if self.year:
                        date_parts = date_part.split('/')
                        if len(date_parts) == 3:
                            y_part = date_parts[2]
                            target_short = self.year[-2:]
                            target_long = self.year
                            if y_part != target_short and y_part != target_long:
                                print(f"Warning: Date {date_part} in header does not match target year {self.year}")

        print(f"Loaded {len(self.date_map)} date columns.")

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
        # Skip header row (0), data starts at row 2 (index 1)
        for i, row in enumerate(rows):
            if i < 1: continue
            
            # Get Column D (Index 3) - Country name (may be hyperlinked)
            cells = row.findall('table:table-cell', self.ns)
            col_idx = 0
            country_name = None
            COUNTRY_COL = 3  # Column D
            
            for cell in cells:
                repeated = cell.attrib.get(f'{{{self.ns["table"]}}}number-columns-repeated')
                count = int(repeated) if repeated else 1
                
                # Check if this cell range covers index 3 (Column D)
                if col_idx <= COUNTRY_COL < col_idx + count:
                    text_node = cell.find('text:p', self.ns)
                    if text_node is not None:
                        # Check for hyperlink first
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

    def get_cell_node(self, row_idx, col_idx):
        """Finds or creates the specific cell node. 
        Note: ODS XML is compressed. Modifying it requires handling 'number-columns-repeated'.
        This version properly splits repeated cells without contaminating other cells.
        """
        import copy
        
        rows = self.table.findall('table:table-row', self.ns)
        if row_idx >= len(rows): return None
        
        row = rows[row_idx]
        cells = row.findall('table:table-cell', self.ns)
        
        current_col = 0
        for i, cell in enumerate(cells):
            repeated = cell.attrib.get(f'{{{self.ns["table"]}}}number-columns-repeated')
            count = int(repeated) if repeated else 1
            
            if current_col <= col_idx < current_col + count:
                # Target is in this group.
                if count > 1:
                    # We MUST split this cell to edit just one.
                    # [Pre_Range] [Target] [Post_Range]
                    
                    pre_count = col_idx - current_col
                    post_count = (current_col + count) - (col_idx + 1)
                    
                    # Make new cells - use deep copy for attributes only, NOT children
                    new_cells_list = []
                    base_attribs = dict(cell.attrib)
                    if f'{{{self.ns["table"]}}}number-columns-repeated' in base_attribs:
                        del base_attribs[f'{{{self.ns["table"]}}}number-columns-repeated']
                    
                    # Pre - empty cells (no content, just structure)
                    if pre_count > 0:
                        pre_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell', base_attribs)
                        pre_cell.attrib[f'{{{self.ns["table"]}}}number-columns-repeated'] = str(pre_count)
                        new_cells_list.append(pre_cell)
                        
                    # Target (The one we return) - fresh empty cell
                    target_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell', base_attribs)
                    new_cells_list.append(target_cell)
                    
                    # Post - empty cells
                    if post_count > 0:
                        post_cell = ET.Element(f'{{{self.ns["table"]}}}table-cell', base_attribs)
                        post_cell.attrib[f'{{{self.ns["table"]}}}number-columns-repeated'] = str(post_count)
                        new_cells_list.append(post_cell)

                    # Replace in parent
                    row.remove(cell)
                    for nc in reversed(new_cells_list): # insert in order
                        row.insert(i, nc)
                    
                    return target_cell
                else:
                    return cell
            
            current_col += count
        return None

    def update_cell_text(self, row_idx, col_idx, text_val, style_name=None):
        cell = self.get_cell_node(row_idx, col_idx)
        if cell is not None:
            # Remove existing text-p
            for child in list(cell):
                if child.tag == f'{{{self.ns["text"]}}}p':
                    cell.remove(child)
            
            p = ET.Element(f'{{{self.ns["text"]}}}p')
            p.text = text_val
            cell.append(p)
            
            # Set type to string
            cell.attrib[f'{{{self.ns["office"]}}}value-type'] = "string"
            
            if style_name:
                cell.attrib[f'{{{self.ns["table"]}}}style-name'] = style_name

    def ensure_amber_style(self):
        """Creates or returns the 'AmberHoliday' cell style with amber background and 6pt font."""
        auto_styles = self.root.find('office:automatic-styles', self.ns)
        if auto_styles is None:
            auto_styles = ET.SubElement(self.root, f'{{{self.ns["office"]}}}automatic-styles')
        
        # Look for existing
        style_name = "AmberHoliday"
        found = False
        for style in auto_styles.findall('style:style', self.ns):
            if style.get(f'{{{self.ns["style"]}}}name') == style_name:
                found = True
                break
        
        if not found:
            style = ET.SubElement(auto_styles, f'{{{self.ns["style"]}}}style', {
                f'{{{self.ns["style"]}}}name': style_name,
                f'{{{self.ns["style"]}}}family': "table-cell",
                f'{{{self.ns["style"]}}}parent-style-name': "Default"
            })
            # Cell properties - amber background
            ET.SubElement(style, f'{{{self.ns["style"]}}}table-cell-properties', {
                f'{{{self.ns["fo"]}}}background-color': "#ffbf00"  # Amber
            })
            # Text properties - 6pt font size
            ET.SubElement(style, f'{{{self.ns["style"]}}}text-properties', {
                f'{{{self.ns["fo"]}}}font-size': "6pt"
            })
        return style_name

    def save(self, output_filename="BH_List_Updated.ods"):
        # We need to repack the zip
        temp_zip = output_filename + ".temp_new"
        
        # If overwriting self.filename, we read from self.filename and write to temp, then move.
        # If output is different, we can just write to output (but we need base files from input).
        
        with zipfile.ZipFile(self.filename, 'r') as zin:
            with zipfile.ZipFile(temp_zip, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == "content.xml":
                        zout.writestr(item, ET.tostring(self.root, encoding='utf-8', xml_declaration=True))
                    else:
                        zout.writestr(item, zin.read(item.filename))
        
        if output_filename == self.filename:
            # Overwrite original
            shutil.move(temp_zip, output_filename)
        else:
            # Move temp to output
            shutil.move(temp_zip, output_filename)
            
        print(f"Saved to {output_filename}")

if __name__ == "__main__":
    import sys
    
    print("--- 2026 Bank Holiday Data Aggregation ---")
    year_input = input("Enter target year (e.g. 2026): ").strip()
    
    if not year_input:
        year_input = "2026" # Default
        print(f"No input, defaulting to {year_input}")
    
    # Main Logic: Check status of all countries
    ods = ODSHandler(year=year_input)
    countries = ods.get_countries()
    print(f"Found {len(countries)} countries in the list.")
    
    missing_data_count = 0
    COUNTRY_COL = 3  # Column D - Country name with hyperlink
    print("\n--- Processing Status ---")
    for row_idx, country_name in countries:
        # Check Column D (Index 3) for hyperlink URL
        url_text = ""
        cell = ods.get_cell_node(row_idx, COUNTRY_COL)
        if cell is not None:
            text_node = cell.find('text:p', ods.ns)
            if text_node is not None:
                # Look for hyperlink (text:a element)
                link = text_node.find('text:a', ods.ns)
                if link is not None:
                    url_text = link.attrib.get(f'{{{ods.ns["xlink"]}}}href', '')
        
        status = "OK" if url_text and len(url_text) > 5 else "MISSING URL"
        if status == "MISSING URL":
            missing_data_count += 1
            print(f"MISSING: {country_name}")
    
    print(f"Total Countries: {len(countries)}")
    print(f"Countries with Missing Data (Pending Scan): {missing_data_count}")
    
    # Ask to proceed with scraping
    if missing_data_count < len(countries): # Meaning we have at least ONE url
        proceed = input("\nProceed with scraping identified URLs? (y/n): ").lower()
        if proceed == 'y':
            scraper = HolidayScraper()
            amber_style = ods.ensure_amber_style()
            updates_made = 0
            
            for row_idx, country_name in countries:
                # Get URL from hyperlink in Column D (Index 3)
                url_text = ""
                cell = ods.get_cell_node(row_idx, COUNTRY_COL)
                if cell is not None:
                    text_node = cell.find('text:p', ods.ns)
                    if text_node is not None:
                        link = text_node.find('text:a', ods.ns)
                        if link is not None:
                            url_text = link.attrib.get(f'{{{ods.ns["xlink"]}}}href', '')
                
                # print(f"DEBUG: {country_name} -> URL: '{url_text}'") # Uncomment for verbose debug
                
                if url_text and "officeholidays.com" in url_text:
                    print(f"Scraping {country_name}...")
                    holidays = scraper.fetch_holidays(url_text)
                    
                    if holidays:
                        print(f"  Found {len(holidays)} holidays.")
                        for date_str, h_name in holidays:
                            if date_str in ods.date_map:
                                col_idx_date = ods.date_map[date_str]
                                # Write to cell
                                ods.update_cell_text(row_idx, col_idx_date, h_name, amber_style)
                            else:
                                # print(f"  Date {date_str} not in headers.")
                                pass
                        updates_made += 1
                        time.sleep(0.5) # Be nice to the server
                    else:
                        print("  No holidays found (or parse error).")
            
            if updates_made > 0:
                ods.save()
                print("Scraping complete and file saved.")
            else:
                print("No updates made.")
        else:
            print("Operation skipped.")
