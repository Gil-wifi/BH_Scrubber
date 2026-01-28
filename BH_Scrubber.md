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
| **Row 1** | Headers: `Supported`, `Region`, `SubRegion`, `Country`, `Website used`, then dates in format `Thu 01/01/26`, `Fri 02/01/26`, etc. |
| **Row 2+** | Country data rows |

### Column Layout

- **Column A (0):** Supported flag
* **Column B (1):** Region
* **Column C (2):** SubRegion  
* **Column D (3):** Country name
* **Column E (4):** Source URL (officeholidays.com)
* **Columns F-NF (5+):** Calendar grid (365 days)

## Instructions

### Phase 1: Source Identification (Official URLs)

For each country listed in Column D starting from **Row 2**, locate the official website containing the public holiday schedule for the year 2026.

* **Action:** Paste the specific URL into **Column E** corresponding to the country's row.
* **Example:** For the UK, the URL is `https://www.officeholidays.com/countries/uk/2026`.
* **Exception:** If no official source exists for a specific country, enter the text `None Available` in Column E.

### Phase 2: Data Population (Calendar Grid)

Scrape or transcribe the holiday data from the identified URLs into the calendar matrix.

* **Calendar Range:** Columns **F** through **NF** (starting from column index 5).
* **Date Format in Headers:** `Thu 01/01/26` (Day abbreviation + DD/MM/YY)
* **Logic:**
  * Locate the cell corresponding to the specific **Country (Row)** and **Date (Column)**.
  * **If a holiday exists:** Input the official name of the holiday in the cell & highlight the cell with an amber color.
  * **If no holiday exists:** Leave the cell entirely blank.

## Scripts

| Script | Purpose |
|--------|---------|
| `bh_Scrubber.py` | Main scraper - fetches holiday data from officeholidays.com |
| `populate_urls.py` | Generates URLs for all countries |
| `fix_404_countries.py` | Fixes URL slugs for countries with 404 errors |

## Definition of Done

The task is considered complete when:

1. Every country in Column D has either a URL or "None Available" in Column E starting from Row 2.
2. All known 2026 public holidays for these countries are mapped accurately by name into the calendar grid (Columns F–NF) starting from Row 2.
