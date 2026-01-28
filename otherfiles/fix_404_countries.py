#!/usr/bin/env python3
"""Fix script for countries with 404 errors due to URL slug mismatches."""
from bh_Scrubber import ODSHandler, HolidayScraper
import time

# Mapping: country name in ODS -> correct officeholidays.com slug
# Found by checking the actual officeholidays.com website structure
FIXES = {
    # EMEA
    "Democratic Republic of the Congo": "congo-democratic-republic",
    "Eswatini": "swaziland",
    "North Macedonia": "macedonia",
    "Republic of the Congo": "congo",
    "United Arab Emirates": "uae",
    # Americas
    "United States of America": "usa",
    # Asia
    "Democratic Republic of Korea": "north-korea",
    "Republic of Korea": "south-korea",
    "Macao": "macau",
    "Bagladesh": "bangladesh",  # Typo in original data
    "Brunei Darussalam": "brunei",
    "Laos": "laos",  # Actually https://www.officeholidays.com/countries/laos
    "Timor-Leste": "east-timor",
    "Viet Nam": "vietnam",
    # Oceania
    "Marshal Islands": "marshall-islands",  # Typo in original data
    "Micronesia": "micronesia",  # May not exist
}

# Countries that likely have no officeholidays.com page (skip)
SKIP = {"Saint Helena", "Nauru", "Niue", "Palau"}

def main():
    # Load the UPDATED file (which has existing data)
    ods = ODSHandler(filename="BH_List_Updated.ods", year=2026)
    scraper = HolidayScraper()
    amber_style = ods.ensure_amber_style()
    
    URL_COL = 4
    
    countries = ods.get_countries()
    updates = 0
    
    for row_idx, country_name in countries:
        if country_name in SKIP:
            print(f"Skipping {country_name} (no known source)")
            continue
            
        if country_name in FIXES:
            slug = FIXES[country_name]
            url = f"https://www.officeholidays.com/countries/{slug}/2026"
            print(f"Fixing {country_name} with URL: {url}")
            
            # Update URL in Column E
            ods.update_cell_text(row_idx, URL_COL, url)
            
            # Scrape
            holidays = scraper.fetch_holidays(url)
            if holidays:
                print(f"  Found {len(holidays)} holidays.")
                for date_str, h_name in holidays:
                    if date_str in ods.date_map:
                        col_idx = ods.date_map[date_str]
                        ods.update_cell_text(row_idx, col_idx, h_name, amber_style)
                updates += 1
                time.sleep(0.3)
            else:
                print(f"  No holidays found for {country_name}.")
    
    if updates > 0:
        # Save back to BH_List.ods (overwrite original)
        ods.save("BH_List.ods")
        print(f"\nFixed {updates} countries. Saved to BH_List.ods")
    else:
        print("No updates made.")

if __name__ == "__main__":
    main()
