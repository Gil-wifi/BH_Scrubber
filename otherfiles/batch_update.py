from bh_Scrubber import ODSHandler

def main():
    ods = ODSHandler()
    
    # Define Batch 1 Data
    # Format: Country Name, URL, [(Date DD/MM/YY, Holiday Name), ...]
    batch_data = [
        {
            "country": "Afghanistan",
            "url": "None Available", 
            "holidays": [
                ("15/02/26", "Liberation Day"),
                ("18/02/26", "Awal Ramadan (Start of Ramadan)"), # Approx
                ("19/02/26", "Awal Ramadan (Start of Ramadan)"), # Alt date usually not marked both? Let's mark 18th as primary
                ("20/03/26", "Eid al-Fitr"),
                ("21/03/26", "Eid al-Fitr Holiday"),
                ("22/03/26", "Eid al-Fitr Holiday"),
                ("28/04/26", "Victory Day"),
                ("01/05/26", "Labour Day"),
                ("26/05/26", "Arafat Day"),
                ("27/05/26", "Eid al-Adha Holiday"),
                ("28/05/26", "Eid al-Adha Holiday"),
                ("29/05/26", "Eid al-Adha Holiday"),
                ("19/08/26", "Independence Day (Jeshen)"),
                ("25/08/26", "Mawleed al-Nabi"),
                ("31/08/26", "Anniversary of Withdrawal")
            ]
        },
        {
            "country": "Albania",
            "url": "https://www.bankofalbania.org/",
            "holidays": [
                ("01/01/26", "New Year's Day"),
                ("02/01/26", "New Year's Day"),
                ("14/03/26", "Summer Day"),
                ("16/03/26", "Summer Day (Observed)"), # 14 is Sat
                ("20/03/26", "Eid al-Fitr"), # Note: weekend rule might apply if 20 is Fri (it is).
                ("22/03/26", "Nevruz Day"),
                ("23/03/26", "Nevruz Day (Observed)"), # 22 is Sun
                ("05/04/26", "Catholic Easter"),
                ("06/04/26", "Catholic Easter Monday"), # Often observed? Bank says "closed on April 6".
                ("12/04/26", "Orthodox Easter"),
                ("13/04/26", "Orthodox Easter Monday"), # Bank says "closed on April 13".
                ("01/05/26", "International Worker's Day"),
                ("27/05/26", "Eid al-Adha"),
                ("05/09/26", "Mother Teresa Day"),
                ("07/09/26", "Mother Teresa Day (Observed)"), # 5 is Sat
                ("22/11/26", "Alphabet Day"),
                ("23/11/26", "Alphabet Day (Observed)"), # 22 is Sun
                ("28/11/26", "Flag and Independence Day"),
                ("30/11/26", "Flag/Independence Day (Observed)"), # 28 is Sat. 
                ("29/11/26", "Liberation Day"),
                ("01/12/26", "Liberation Day (Observed)"), # 29 is Sun.
                ("08/12/26", "National Youth Day"),
                ("25/12/26", "Christmas Day")
            ]
        },
        {
            "country": "Algeria",
            "url": "https://www.mfa.gov.dz/", # Ministry of Foreign Affairs
            "holidays": [
                ("01/01/26", "New Year's Day"),
                ("12/01/26", "Amazigh New Year"),
                ("22/02/26", "Day of Fraternity and Unity"),
                ("20/03/26", "Eid al-Fitr"), # Approx
                ("21/03/26", "Eid al-Fitr Holiday"),
                ("22/03/26", "Eid al-Fitr Holiday"),
                ("01/05/26", "Labour Day"),
                ("27/05/26", "Eid al-Adha"), # Approx
                ("28/05/26", "Eid al-Adha Holiday"),
                ("29/05/26", "Eid al-Adha Holiday"),
                ("16/06/26", "Islamic New Year"), # Approx
                ("25/06/26", "Ashura"), # Approx
                ("05/07/26", "Independence Day"),
                ("03/09/26", "Prophet Muhammad's Birthday"), # Approx
                ("01/11/26", "Anniversary of the Revolution")
            ]
        },
        {
            "country": "Angola",
            "url": "https://www.siac.gv.ao/", # Official Portal
            "holidays": [
                ("01/01/26", "New Year's Day"),
                ("04/02/26", "Repatriation Day"), # "Armed Struggle"
                ("17/02/26", "Carnival"),
                ("08/03/26", "International Women's Day"),
                ("23/03/26", "Southern Africa Liberation Day"),
                ("03/04/26", "Good Friday"),
                ("04/04/26", "Peace Day"),
                ("01/05/26", "Labour Day"),
                ("17/09/26", "National Heroes' Day"),
                ("02/11/26", "All Souls' Day"),
                ("11/11/26", "Independence Day"),
                ("25/12/26", "Christmas Day")
            ]
        }
    ]
    
    # Process batches
    countries_map = {name: idx for idx, name in ods.get_countries()}
    amber_style = ods.ensure_amber_style()

    for item in batch_data:
        c_name = item["country"]
        if c_name not in countries_map:
            print(f"Country {c_name} not found in ODS.")
            continue
            
        row_idx = countries_map[c_name]
        print(f"Updating {c_name} (Row {row_idx})...")
        
        # Update URL (Column C -> Index 2)
        # Assuming index 2 is Column C. Row indices are 0-based in list, but methods use them directly.
        ods.update_cell_text(row_idx, 2, item["url"])
        
        # Update Holidays
        for date_str, holiday_name in item["holidays"]:
            if date_str in ods.date_map:
                col_idx = ods.date_map[date_str]
                ods.update_cell_text(row_idx, col_idx, holiday_name, amber_style)
            else:
                print(f"  Date {date_str} not found in header.")

    ods.save("BH_List.ods") # Overwrite

if __name__ == "__main__":
    main()
