# MRBR Report Automation

**Status: Production**

Python automation for processing and preparing MRBR reports for the Purchasing team.

## Overview

This automation processes an MRBR Excel report exported by an administrator to a designated server location. It applies predefined Purchasing filtering and cleanup rules, prepares the report for review, and saves the processed workbook to the designated IMAC server output location.

The automation reduces repetitive manual report preparation and provides a consistent output format for the Purchasing process.

## Business Problem

MRBR reports require manual preparation before they can be used by the Purchasing team. The process involves reviewing multiple report worksheets, filtering relevant plants, applying specific business rules, and preparing additional fields for Purchasing follow-up.

Performing these steps manually is repetitive and can lead to inconsistent report preparation.

## Solution

The Python automation:

* Reads the `SITE MRBR`, `IPO MRBR`, and `IMAC MRBR` worksheets.
* Filters records to the required Purchasing plants.
* Applies the IPO Company Code filtering rule.
* Adds Purchasing review fields.
* Creates a cleaned Excel workbook.
* Saves the processed report to the designated IMAC server location.

## Workflow

![Workflow](docs/diagrams/workflow.png)

**Process:**

`MRBR Export → Server Input → Python Processing → Filtering & Cleanup → Processed MRBR → Server Output`

## Architecture

![Architecture](docs/diagrams/architecture.png)

The automation uses Python and pandas to process the exported Excel workbook. SAP is outside the direct automation boundary; the Python process starts after the MRBR report has been exported.

## Technologies

* Python
* pandas
* Microsoft Excel
* openpyxl

## Key Features

* Multi-sheet MRBR processing
* Plant-based filtering
* IPO Company Code filtering
* Automated Purchasing review columns
* Automated output workbook generation
* Server-based input and output workflow

## Input & Output

**Input:**
MRBR Excel report exported by an administrator to the designated server location.

**Output:**
Processed MRBR Excel workbook saved to the designated IMAC server output folder.

## My Role

I designed and developed the automation based on the Purchasing workflow. My responsibilities included defining the processing logic and business rules, developing the Python solution, testing the processed output, and maintaining and improving the automation.

## Limitations

* The automation depends on the expected MRBR Excel worksheet and column structure.
* Plant and business-rule values are currently defined within the Python implementation.
* The automation does not directly execute SAP transactions or control SAP GUI.

## Future Improvements

Potential improvements include externalizing business rules and configuration, adding stronger input validation, and introducing structured logging and automated testing.

## Disclaimer

This repository contains a portfolio representation of a production business automation. Company-specific data, credentials, and confidential information are not included. Sample data is provided for demonstration purposes.
