# Project: 2026 Bank Holiday Data Aggregation

## Overview

The objective of this project is to populate a master dataset with official 2026 public holiday information for a specific list of countries. The source file is `BH_List.ods`.

## Input Data

* **File:** `BH_List.ods`
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

Holiday data is populated into the calendar matrix with **6pt font size** and **amber highlighting**.

* **Calendar Range:** Columns **E** through **ND** (starting from column index 4)
* **Date Format:** `Thu 01/01/26` (Day abbreviation + DD/MM/YY)
* **Styling:** Amber background (#ffbf00), 6pt font

## Scripts

| Script              | Purpose                                             |
|---------------------|-----------------------------------------------------|
| `bh_Scrubber.py`    | Main scraper (v1.2) - reads hyperlinks from Col D   |
| `populate_urls.py`  | Generates URLs for all countries                    |
| `fix_404_countries.py` | Fixes URL slugs for 404 errors                   |

## Version History

* **v1.4** - Bug fixes, verified holiday population working correctly
* **v1.2** - Hyperlinks on country names, 6pt font for holidays, calendar starts Col E
* **v1.0** - Initial release

## Author

Gil Alowaifi
