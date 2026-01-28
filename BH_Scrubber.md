# Project: 2026 Bank Holiday Data Aggregation

## Overview

The objective of this project is to populate a master dataset with official 2026 public holiday information for a specific list of countries. The source file for this task is `BH_List.ods`.

## Input Data

* **File:** `BH_List.ods`
* **Target Countries:** Listed in **Column D** starting from **Row 2**.

## File Structure

The ODS file has an optimized single-header structure:

| Row | Content |
|-----|---------|
| **Row 1** | Headers: `Supported`, `Region`, `SubRegion`, `Country`, then dates in format `Thu 01/01/26` |
| **Row 2+** | Country data rows |

### Column Layout

* **Column A (0):** Supported flag
* **Column B (1):** Region
* **Column C (2):** SubRegion  
* **Column D (3):** Country name **with hyperlink to source URL**
* **Columns E-NE (4+):** Calendar grid (365 days)

## Instructions

### Phase 1: Source Identification (Official URLs)

For each country in Column D, the source URL is stored as a **hyperlink on the country name**.

* **Format:** Country name is clickable and links to `https://www.officeholidays.com/countries/{country}/2026`
* **Exception:** If no source exists, country name has no hyperlink

### Phase 2: Data Population (Calendar Grid)

Holiday data is populated into the calendar matrix.

* **Calendar Range:** Columns **E** through **NE** (starting from column index 4)
* **Date Format in Headers:** `Thu 01/01/26` (Day abbreviation + DD/MM/YY)
* **Logic:**
  * Locate the cell corresponding to the specific **Country (Row)** and **Date (Column)**
  * **If a holiday exists:** Input the official name of the holiday in the cell & highlight amber
  * **If no holiday exists:** Leave the cell blank

## Scripts

| Script | Purpose |
|--------|---------|
| `bh_Scrubber.py` | Main scraper (v1.2) - reads hyperlinks from Column D |
| `populate_urls.py` | Generates URLs for all countries |
| `fix_404_countries.py` | Fixes URL slugs for countries with 404 errors |

## Version History

* **v1.2** - Hyperlinks on country names, removed URL column E
* **v1.0** - Initial release

## Definition of Done

The task is considered complete when:

1. Every country in Column D has a hyperlink to its source URL (or no link if unavailable)
2. All known 2026 public holidays are mapped into the calendar grid starting from Row 2
