# MRBR Report

This Python script automates the processing of MRBR (Blocked Invoice) reports exported from SAP.  
It consolidates data from multiple sheets, applies filters, and generates a cleaned Excel output for easier review and approval.

---

## Features
- Reads MRBR data from multiple sheets (`IPO MRBR`, `SITE MRBR`, `IMAC MRBR`).
- Filters out unwanted records (e.g., Company Code `1803`).
- Keeps only relevant plants (`HU07`, `HU08`, `IN07`, `IT08`, `PL01`, `SG02`, `VN01`).
- Adds helper columns: `BUYER Comment`, `Elaine approval`, `SP Update`.
- Saves the processed report to a structured server location with today’s date.

---

## File Structure
- `main.py` → Main script for processing MRBR reports.
- Output file → Saved automatically to:
  - `//sgsind0nsifsv01a/IMAC Data/.../{year}/MRBR {today}.xlsx`

---

## How It Works
1. User provides an input Excel file name when prompted.
2. Script reads data from:
   - **SITE MRBR**
   - **IPO MRBR**
   - **IMAC MRBR**
3. Filters are applied:
   - Excludes CoCd = `1803`
   - Keeps only selected plants
4. Cleaned data is written to a new Excel file with the same sheet structure.
