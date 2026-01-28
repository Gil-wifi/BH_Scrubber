# BH_Scrubber (Bank Holiday Scraper)

![Version](https://img.shields.io/badge/version-2.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.x-green.svg)

**BH_Scrubber** is a specialized Python automation tool designed to scrape, aggregate, and classify global public holiday data from [OfficeHolidays.com](https://www.officeholidays.com) into a master `.ods` spreadsheet.

## 🎯 Motivation & Scope

This project was built to solve a critical **Workforce Optimization (WFO)** challenge: efficiently tracking bank holidays across a multinational organization.

Historically, scheduling conflicts arose because all holidays were treated equally. Colleagues were often unaware that a holiday in a specific country might be **Regional** (affecting only a subset of the workforce) rather than **National** (affecting everyone).

**BH_Scrubber** addresses this by:

1. **Distinguishing Holiday Types:** It clearly identifies and visually isolates **National** holidays (Amber) from **Regional** ones (Pink).
2. **Visualizing Support:** It highlights supported countries (Green) and greys out unsupported ones (Grey), providing an instant visual status map for planners.
3. **Automating Updates:** It eliminates manual data entry errors by scraping the latest data directly from a trusted source.

## ✨ Key Features (v2.0)

* **Intelligent Scraping:** Parses `officeholidays.com` to extract dates, names, and holiday types.
* **Smart Classification:**
  * **National Holidays:** Flagged and styled in **Amber** (#ffbf00).
  * **Regional Holidays:** Detected and styled in **Pink** (#ffb6c1).
* **Visual Status Indicators:**
  * ✅ **Supported Countries:** Rows colored **Light Green**.
  * ❌ **Unsupported/Excluded:** Rows colored **Medium Grey** (Visual Blackout).
* **ODS Integration:** Directly modifies OpenDocument Spreadsheet (ODS) XML for high-performance updates without breaking existing formulas or structures.
* **Timestamped Output:** Generates unique output files (e.g., `BH_List_20260128.ods`) to preserve history.

## 🚀 Usage

1. **Prepare the Input File:**
    * Ensure `BH_List.ods` is present.
    * To support a country, mark **Column A** as `Yes`. To exclude it, mark `No`.
    * Ensure the country name in **Column D** has a hyperlink to its `officeholidays.com` page.

2. **Run the Scraper:**

    ```bash
    python3 bh_Scrubber.py
    ```

3. **Follow Prompts:**
    * Enter the target year (e.g., `2026`).
    * Review the country count and integrity check.
    * Confirm to start the scraping process.

4. **View Results:**
    * The tool saves a new file: `BH_List_YYYYMMDD.ods`.
    * Open it in LibreOffice Calc or Excel to view the populated and styled calendar.

## 🛠 Project Structure

* `bh_Scrubber.py`: Main script containing scraping and ODS manipulation logic.
* `BH_List.ods`: Template spreadsheet.
* `BH_Scrubber.md`: Detailed technical documentation.

## 👤 Author

**Gil Alowaifi**

---
*Disclaimer: This tool is for internal workforce planning purposes. Data is sourced from public websites.*
