# MRBR Report Automation

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
```

Records belonging to other plants are excluded from the processed output.

---

### IPO Company Code Filtering

For the `IPO MRBR` worksheet, records containing Company Code `1803` are excluded.

This rule is applied specifically to the IPO dataset.

---

### Purchasing Review Columns

The automation adds the following columns to the processed datasets:

```text
BUYER Comment
Elaine approval
SP Update
```

These fields provide space for subsequent Purchasing review and follow-up activities.

---

### Automated Output Generation

The processed data is written into a new Excel workbook.

The output filename uses the current date in the following format:

```text
MRBR YYMMDD.xlsx
```

The generated workbook is saved to the designated server location.

---

## Technologies Used

| Category | Technology |
|---|---|
| Programming Language | Python |
| Data Processing | pandas |
| Input | Microsoft Excel |
| Output | Microsoft Excel |
| Excel Processing | pandas Excel I/O |
| File Handling | Python |

---

# Workflow

The MRBR automation follows the workflow below.

![MRBR Automation Workflow](docs/diagrams/workflow.png)

### Workflow Steps

**1. Input MRBR Workbook**

The user provides the MRBR Excel workbook used for processing.

**2. Read MRBR Worksheets**

The automation reads:

- `SITE MRBR`
- `IPO MRBR`
- `IMAC MRBR`

**3. Apply Plant Filtering**

The automation retains only the required plants:

```text
HU07
HU08
IN07
IT08
PL01
SG02
VN01
```

**4. Apply IPO Company Code Rule**

For the IPO MRBR data, records containing Company Code `1803` are excluded.

**5. Add Purchasing Review Columns**

The automation adds:

```text
BUYER Comment
Elaine approval
SP Update
```

**6. Generate Output Workbook**

The processed datasets are written into a new Excel workbook.

**7. Save Output**

The processed workbook is saved using the required date-based filename.

---

# Architecture

The current implementation uses Python and pandas to read, filter, transform, and generate the MRBR Excel report.

![MRBR Automation Architecture](docs/diagrams/architecture.png)

### Technical Components

| Component | Responsibility |
|---|---|
| `MRBR.py` | Main Python automation script |
| `pandas` | Read, filter, transform, and write Excel data |
| Excel Input | Source MRBR workbook |
| Business Rules | Apply plant and Company Code filtering |
| Data Transformation | Add Purchasing review fields |
| `ExcelWriter` | Generate the processed Excel workbook |
| File Handling | Manage input/output paths and date-based filenames |

The current implementation uses a straightforward data-processing workflow rather than direct SAP GUI automation.

---

# Before / After Example

A sample Excel workbook is included to demonstrate the transformation performed by the automation.

### Sample Workbook

[Open the MRBR Before / After Sample](MRBR_Before_After_Sample.xlsx)

The sample demonstrates the report before and after automated processing, including:

- Plant filtering.
- IPO Company Code filtering.
- Addition of Purchasing review columns.
- Generation of the processed workbook.

The sample is provided for demonstration purposes and does not represent live production data.

---

# Business Rules

## Plant Filtering

The automation retains only the following plants:

```text
HU07
HU08
IN07
IT08
PL01
SG02
VN01
```

---

## IPO Company Code Filtering

For the `IPO MRBR` worksheet, records containing:

```text
Company Code 1803
```

are excluded.

---

## Purchasing Review Fields

The following columns are added during processing:

```text
BUYER Comment
Elaine approval
SP Update
```

---

# Business Impact

The automation helps the Purchasing team by:

- Reducing repetitive Excel report preparation.
- Standardizing MRBR filtering.
- Applying consistent business rules.
- Automatically adding required Purchasing review fields.
- Generating the processed workbook automatically.
- Applying a consistent output naming convention.
- Providing a repeatable report-preparation workflow.

No percentage-based efficiency improvement is claimed because formally measured time-saving data is not available.

---

# My Role

I was responsible for:

- Understanding the MRBR reporting process used by the Purchasing team.
- Identifying repetitive report-preparation activities.
- Designing the automation workflow.
- Developing the Python automation.
- Implementing the plant filtering logic.
- Implementing the IPO Company Code filtering rule.
- Using pandas for Excel data processing.
- Automating the Excel output generation.
- Implementing the date-based output filename.
- Testing the processed report.
- Maintaining and improving the automation.

---

# Technical Implementation

The main automation is contained in:

```text
MRBR.py
```

The implementation uses pandas to:

- Read Excel worksheets.
- Convert required data fields for filtering.
- Filter records based on defined business rules.
- Add Purchasing review columns.
- Write processed DataFrames into a new Excel workbook.

The script also handles:

- Input file path construction.
- Output path construction.
- Date-based output naming.
- Excel workbook generation.

The processing is performed separately for each MRBR worksheet before the results are written to the output workbook.

---

# Current Limitations

- The automation depends on the expected MRBR worksheet names.
- The automation depends on the expected structure of the MRBR Excel report.
- The required plant list is currently defined in the Python script.
- The IPO Company Code filtering rule is currently defined in the Python script.
- The source MRBR workbook must be available before processing.
- The automation currently processes Excel data after the MRBR report has been exported from SAP.
- The automation does not directly execute SAP transactions.
- The current implementation is designed around the existing MRBR report structure.

---

# Lessons Learned

This project provided practical experience in applying Python to a real Purchasing reporting process.

Key lessons include:

- Repetitive Excel preparation can be converted into a repeatable Python workflow.
- Business rules should be explicitly implemented rather than manually applied each time.
- pandas provides a practical approach for processing structured Excel data.
- Automating report preparation can reduce repetitive manual handling even without direct SAP integration.
- Separating input, processing, and output steps makes the automation easier to understand and maintain.
- Output naming and file handling are also important parts of a business automation workflow.

---

# Roadmap

Potential future improvements include:

### Configuration

Move business rules such as the plant list and Company Code exclusions into a separate configuration file.

### Input Validation

Validate the input workbook before processing to detect:

- Missing worksheets.
- Missing required columns.
- Unexpected report structures.
- Invalid input files.

### Error Handling

Improve error handling and provide clearer feedback when:

- The source file cannot be found.
- Expected worksheets are missing.
- Required columns are missing.
- The output location is unavailable.
- The report structure has changed.

### Logging

Introduce structured logging to record:

- Processing start and completion.
- Input file.
- Output file.
- Number of records processed.
- Filtering results.
- Processing errors.

### Modular Design

Separate the processing logic into reusable modules for:

- Input handling.
- Validation.
- Business rules.
- Data transformation.
- Output generation.

### User Interface

Provide a simple interface for:

- Selecting the input MRBR workbook.
- Starting the processing.
- Viewing the processing result.
- Selecting the output location.

---

# Engineering Skills Demonstrated

- Python
- pandas
- Excel Automation
- Multi-Sheet Excel Processing
- Data Filtering
- Business Rule Implementation
- Data Transformation
- Automated Report Generation
- File and Path Handling
- Business Process Automation
- Troubleshooting
- Maintenance

---

# Project Information

| Item | Details |
|---|---|
| Project Status | Production |
| Project Type | Business Process Automation |
| Primary Function | MRBR Report Processing |
| Programming Language | Python |
| Data Processing | pandas |
| Input | SAP MRBR Excel Workbook |
| Output | Processed Excel Workbook |
| Primary Users | Purchasing Team |
