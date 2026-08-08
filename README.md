# MRBR Automation

Python-based report automation developed to process SAP MRBR (Blocked Invoice) reports and prepare standardized Excel output for Purchasing review.

---

## Overview

MRBR is used to identify blocked invoices that require Purchasing review and follow-up.

The manual process involved working with multiple MRBR report files, applying plant and company-code filtering, and adding the required review fields before producing the final working file.

This automation standardizes those steps using Python and pandas.

---

## Business Problem

The MRBR process required repetitive Excel preparation, including:

- Processing multiple MRBR report types
- Filtering records by required plants
- Applying specific business rules
- Adding Purchasing review columns
- Preparing a standardized output file
- Renaming the final file using the required date format

Performing these steps manually increases repetitive work and creates opportunities for inconsistent filtering or formatting.

---

## Solution

The automation reads the required MRBR Excel reports, processes the data according to predefined business rules, and generates a standardized Excel output for Purchasing review.

### Process

1. Read the MRBR input reports
2. Process the required MRBR report types
3. Filter records based on required plants
4. Apply the IPO company-code exclusion rule
5. Add required Purchasing review columns
6. Generate the processed Excel output
7. Save the output using the required naming convention

---

## Supported MRBR Reports

The automation processes the following MRBR reports:

- `SITE MRBR`
- `IPO MRBR`
- `IMAC MRBR`

---

## Business Rules

### Plant Filtering

The following plants are included in the processing:

```text
HU07
HU08
IN07
IT08
PL01
SG02
VN01
