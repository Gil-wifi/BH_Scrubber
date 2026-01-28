# Project: 2026 Bank Holiday Data Aggregation

## Overview

The objective of this project is to populate a master dataset with official 2026 public holiday information for a specific list of countries. The source file for this task is `BH_List.ods`.

## Input Data

* **File:** `BH_List.ods`
* **Target Countries:** Listed in **Column D**.

## Instructions

### Phase 1: Source Identification (Official URLs)

For each country listed in Column D starting from cell D6, locate the official **governmental** website containing the public holiday schedule for the year 2026.

* **Action:** Paste the specific URL into **Column E** corresponding to the country's row starting from row 6.
* **Example:** For the UK, the URL is `https://www.gov.uk/bank-holidays`.
* **Exception:** If no official governmental source exists for a specific country, enter the text `None Available` in Column E.

### Phase 2: Data Population (Calendar Grid)

Scrape or transcribe the holiday data from the identified URLs into the calendar matrix.

* **Calendar Range:** Columns **F** through **NF** starting from column F6.
* **Logic:**
  * Locate the cell corresponding to the specific **Country (Row)** and **Date (Column)**.
  * **If a holiday exists:** Input the official name of the holiday in the cell & highlight the cell with an amber color.
  * **If no holiday exists:** Leave the cell entirely blank.

## Definition of Done

The task is considered complete when:

1. Every country in Column D has either a URL or "None Available" in Column E starting from row 6.
2. All known 2026 public holidays for these countries are mapped accurately by name into the calendar grid (Columns F–NF) starting from column F6.
