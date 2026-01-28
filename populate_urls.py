from bh_Scrubber import ODSHandler
import re

def slugify(name):
    """Converts country name to officeholidays slug."""
    # "Bosnia & Herzegovina" -> "bosnia-and-herzegovina"
    name = name.strip().lower()
    name = name.replace("&", "and")
    name = name.replace(" ", "-")
    name = name.replace("'", "") # Cote d'Ivoire -> cote-divoire? usually just remove chars
    # Special cases might exist, but this covers 90%
    return name

def main():
    ods = ODSHandler()
    countries = ods.get_countries()
    
    URL_COL = 4  # Column E - Updated for new structure
    
    print(f"Checking URLs for {len(countries)} countries...")
    updates = 0
    
    for row_idx, country_name in countries:
        # Check if URL is present (Column E / Index 4)
        cell = ods.get_cell_node(row_idx, URL_COL)
        current_url = ""
        if cell is not None:
            text_node = cell.find('text:p', ods.ns)
            if text_node is not None:
                current_url = text_node.text
        
        # Force update or check if valid
        # if not current_url or current_url == "None Available" or len(current_url) < 10:
        if True: # Force update to fix any bad slugs
            # Generate URL
            slug = slugify(country_name)
            # Correcting specific known mismatches if any
            if country_name == "Democratic Republic of the Congo": slug = "democratic-republic-of-the-congo" # Check this?
            
            new_url = f"https://www.officeholidays.com/countries/{slug}/2026"
            
            print(f"Setting {country_name} -> {new_url}")
            ods.update_cell_text(row_idx, URL_COL, new_url)
            updates += 1
            
    if updates > 0:
        ods.save("BH_List.ods")
        print(f"Updated {updates} countries with URLs.")
    else:
        print("No updates needed.")

if __name__ == "__main__":
    main()
