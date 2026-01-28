# Project: 2026 Bank Holiday Data Aggregation

## Overview

The objective of this project is to populate a master dataset with official 2026 public holiday information for a specific list of countries. The source file is `BH_List.ods`.

## Motivation

This tool was created to simplify the WFO (Workforce Optimization) task of tracking holidays across different countries. Specifically, it distinguishes between **National/Public** holidays (which affect everyone) and **Regional** holidays. Previously, treating all holidays as "National" caused confusion in work programming and scheduling, as colleagues were often unaware of these distinctions. By visualizing these differences, we ensure accurate schedule planning.

## Input Data

* **File:** `country_List.ods` (Template)
* **Output:** `BH_List_YYYYMMDD.ods` (Generated daily)
* **Target Countries:** Listed in **Column D** starting from **Row 2**

## File Structure

The ODS file has an optimized single-header structure:

| Row       | Content                                                                      |
|-----------|------------------------------------------------------------------------------|
| **Row 1** | Headers: `Supported`, `Region`, `SubRegion`, `Country`, then dates           |
| **Row 2+**| Country data rows                                                            |

### Column Layout

* **Column A (0):** Supported flag
* **Column B (1):** Region
* **Column C (2):** SubRegion  
* **Column D (3):** Country name **with hyperlink to source URL**
* **Columns E-ND (4+):** Calendar grid (365 days) - format: `Thu 01/01/26`

## Instructions

### Phase 1: Source Identification

Source URLs are stored as **hyperlinks on the country name** in Column D.

* **Format:** `https://www.officeholidays.com/countries/{country}/2026`

### Phase 2: Data Population

Holiday data is populated into the calendar matrix.

* **Calendar Range:** Generated programmatically starting from Column **E**.
* **Date Format:** `Thu 01/01/26` (Day abbreviation + DD/MM/YY)
* **Calendar Types:**
  * **Solar:** Jan 1st - Dec 31st.
  * **Fiscal:** Custom start date (Jan 1st + Offset Weeks). Spans full year duration.
* **Visual Styling (Output File):**
  * **Row Styling (Columns A-ND):**
    * **Supported Countries (Col A = "Yes"):** Rows are highlighted in **Light Green** to indicate active tracking.
    * **Unsupported Countries (Col A = "No"):** Rows are greyed out in **Medium Grey** to visually exclude them from planning.
  * **Holiday Cells:**
    * **National/Public Holidays:** **Amber** background (#ffbf00) with 6pt font. These are critical for scheduling.
    * **Regional/Local Holidays:** **Light Pink** background (#ffb6c1) with 6pt font. These are informational and may not affect the entire workforce.

## Scripts

| Script              | Purpose                                             |
|---------------------|-----------------------------------------------------|
| `bh_Scrubber.py`    | Main scraper (v1.6) - supports National vs Regional |
| `populate_urls.py`  | Generates URLs for all countries                    |
| `fix_404_countries.py` | Fixes URL slugs for 404 errors                   |

## Known Limitations

* **Future Data Availability:** Sources like `officeholidays.com` may not have published data for future years (e.g., 2027) at the time of scraping. In these cases, the tool will skip the missing year and log a note.

## Version History

* **v2.1** - Added Automatic Calendar Generation, Fiscal Year Support, and Multi-Year Scraping.
* **v2.0** - Implemented Row Styling: Green for 'Yes' (Supported), Grey for 'No' (Unsupported).
* **v1.6** - Added support for Regional (Pink) vs National (Amber) holidays
* **v1.4** - Bug fixes, verified holiday population
* **v1.2** - Hyperlinks on country names, 6pt font, calendar starts Col E
* **v1.0** - Initial release

## Author

Gil Alowaifi
