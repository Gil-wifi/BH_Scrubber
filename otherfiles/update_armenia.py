from bh_Scrubber import ODSHandler

def main():
    ods = ODSHandler()
    
    # Define Armenia Data
    batch_data = [
        {
            "country": "Armenia",
            "url": "https://www.gov.am/en/holidays/", # Generic Gov URL or Wikipedia
            "holidays": [
                ("01/01/26", "New Year's Day"),
                ("02/01/26", "New Year's Day Holiday"),
                ("06/01/26", "Christmas Day"),
                ("28/01/26", "Army Day"),
                ("08/03/26", "International Women's Day"),
                ("06/04/26", "Easter Monday"), # Based on search
                ("24/04/26", "Genocide Remembrance Day"),
                ("01/05/26", "Labor Day"),
                ("09/05/26", "Victory and Peace Day"),
                ("28/05/26", "First Republic Day"),
                ("05/07/26", "Constitution Day"),
                ("21/09/26", "Independence Day"),
                ("31/12/26", "New Year's Eve")
            ]
        }
    ]
    
    countries_map = {name: idx for idx, name in ods.get_countries()}
    amber_style = ods.ensure_amber_style()

    for item in batch_data:
        c_name = item["country"]
        if c_name not in countries_map:
            print(f"Country {c_name} not found in ODS.")
            continue
            
        row_idx = countries_map[c_name]
        print(f"Updating {c_name} (Row {row_idx})...")
        
        ods.update_cell_text(row_idx, 2, item["url"])
        
        for date_str, holiday_name in item["holidays"]:
            if date_str in ods.date_map:
                col_idx = ods.date_map[date_str]
                ods.update_cell_text(row_idx, col_idx, holiday_name, amber_style)
            else:
                print(f"  Date {date_str} not found.")

    ods.save("BH_List.ods") # Overwrite

if __name__ == "__main__":
    main()
