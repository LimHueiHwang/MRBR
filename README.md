# MRBR Automation

![Production](https://img.shields.io/badge/Status-Production-success)
![Python](https://img.shields.io/badge/Python-Automation-blue)
![pandas](https://img.shields.io/badge/pandas-Data%20Processing-green)
![Excel](https://img.shields.io/badge/Excel-Automation-orange)

> Python-based report automation developed to process SAP MRBR (Blocked Invoice) reports and prepare standardized Excel output for Purchasing review.

---

## Overview

MRBR is used to identify blocked invoices that require Purchasing review and follow-up.

This automation processes an MRBR Excel workbook containing three report worksheets:

- `SITE MRBR`
- `IPO MRBR`
- `IMAC MRBR`

The Python script applies predefined Purchasing business rules, adds the required review fields, and generates a processed Excel workbook for Purchasing review.

**Project Status:** Production

---

## Business Problem

The MRBR report required repetitive Excel preparation before the Purchasing team could perform its review.

The manual process involved:

- Opening the MRBR Excel workbook.
- Processing the required MRBR report types.
- Filtering records based on required plants.
- Applying a specific Company Code exclusion rule.
- Adding Purchasing review columns.
- Preparing the processed workbook.
- Saving the output using the required date-based naming convention.

These repetitive steps created unnecessary manual work and could lead to inconsistent filtering or preparation.

---

## Solution

The automation standardizes the MRBR report-preparation process using Python and pandas.

The script:

1. Reads the MRBR Excel workbook.
2. Loads the `SITE MRBR`, `IPO MRBR`, and `IMAC MRBR` worksheets.
3. Filters the data to the required plants.
4. Excludes records containing Company Code `1803` from the IPO MRBR data.
5. Adds Purchasing review columns.
6. Generates a new Excel workbook containing the processed data.
7. Saves the output using the required date-based filename.

The automation processes the Excel report exported from SAP. It does not directly execute SAP transactions.

---

## Key Features

### Multi-Sheet Excel Processing

The automation processes three MRBR worksheets:

- `SITE MRBR`
- `IPO MRBR`
- `IMAC MRBR`

Each worksheet is processed separately and written into the generated output workbook.

---

### Plant Filtering

The automation retains the following plants:

```text
HU07
HU08
IN07
IT08
PL01
SG02
VN01
