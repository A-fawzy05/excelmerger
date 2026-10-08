# Excel Merger

A desktop application for consolidating multiple Excel workbooks into a single, organized spreadsheet, with intelligent column normalization, key-based record merging, source tracking, and formatting preservation.

Built with **Python, PySide6, and openpyxl**, Excel Merger is designed to simplify spreadsheet consolidation, especially when working with Arabic-language data and Excel files that use inconsistent column names or number formats.

## Table of Contents

* [Overview](#overview)
* [Features](#features)
* [How It Works](#how-it-works)
* [Technology Stack](#technology-stack)
* [Project Structure](#project-structure)
* [Prerequisites](#prerequisites)
* [Installation](#installation)
* [Running the Application](#running-the-application)
* [Using the Application](#using-the-application)
* [Merge Logic and Data Handling](#merge-logic-and-data-handling)
* [Arabic Data Normalization](#arabic-data-normalization)
* [Output Format](#output-format)
* [Troubleshooting](#troubleshooting)
* [Development and Contribution](#development-and-contribution)
* [License](#license)

## Overview

Excel Merger helps users consolidate information from multiple Excel workbooks without manually copying and pasting rows between spreadsheets.

Instead of handling each workbook individually, users can select source files, choose the worksheets to include, and optionally specify key columns to determine which records should be combined.

The application collects the selected data, aligns equivalent columns, combines matching records according to its merge rules, and generates a new Excel workbook.

Special attention is given to Arabic-language spreadsheets, including normalization of common Arabic header variations and Arabic or Persian numerals.

### Why Excel Merger?

Spreadsheet consolidation can become difficult when:

* Information is distributed across several workbooks.
* Different files use slightly different names for the same column.
* The same record appears in multiple files with complementary information.
* Some cells are empty in one file but populated in another.
* Users need to identify which source files contributed to each record.
* Arabic text and numerals are represented inconsistently.

Excel Merger addresses these problems through a desktop interface and a centralized merging process.

## Features

### 1. Multi-Workbook Consolidation

* Load multiple Excel workbooks into a single consolidation workflow.
* Collect data from selected worksheets across the input files.
* Combine data into a new output workbook.
* Avoid manually transferring rows between spreadsheets.

### 2. Key-Based Record Matching

When key columns are selected, the application can group records that share the same key values.

For example, a unique identifier can be used to recognize that two rows refer to the same person or entity, even when they originate from different workbooks.

The application normalizes key values so that common numeric representations, such as `123`, `123.0`, and `١٢٣`, can match.

**Important:** Matching depends on the selected key columns. Without selected keys, records are kept separate rather than being automatically deduplicated.

### 3. Automatic Column Normalization

The application normalizes column headers to recognize certain equivalent spellings.

For example, configured Arabic aliases include:

| Canonical column | Recognized aliases     |
| ---------------- | ---------------------- |
| `الاسم`          | `الإسم`, `اسم`         |
| `الهاتف`         | `رقم الهاتف`, `موبايل` |

This helps align equivalent columns across workbooks without requiring every source file to use precisely the same header spelling.

### 4. Missing-Value Completion

When multiple records match, the application combines their available cell values.

If a column is empty in the existing record and a matching record provides a value, the populated value can fill the missing information.

When conflicting non-empty values exist, the application retains the value encountered first according to its processing order.

### 5. Source Tracking

The generated workbook includes an additional column named `اسم_الملف`.

This column records the source workbook and worksheet associated with each consolidated record, helping users trace information back to its origin.

When a record is assembled from multiple sources, the contributing source labels are combined in the output.

### 6. Excel Formatting Preservation

The application attempts to preserve relevant cell formatting from the source workbook, including:

* Font formatting.
* Cell fills and borders.
* Cell alignment.
* Cell protection settings.
* Number formats.
* Header appearance and column widths.
* Available row-height information.

The generated workbook uses a right-to-left worksheet layout, which is useful for Arabic-language spreadsheets.

### 7. Desktop Graphical Interface

The application uses PySide6 to provide a native desktop interface for interacting with the spreadsheet-merging workflow.

Users can work with Excel files through the graphical application rather than needing to write their own Python scripts.

## How It Works

The merging workflow can be summarized as follows:

1. **Select input workbooks:** Provide the Excel files that contain the data to consolidate.
2. **Select worksheets:** Choose the worksheets to include in the operation.
3. **Configure matching keys:** Optionally specify columns used to identify equivalent records.
4. **Normalize headers and keys:** Apply supported normalization rules to improve matching consistency.
5. **Collect and combine records:** Align columns, group matching records when keys are configured, and fill missing values according to the merge rules.
6. **Track source information:** Associate consolidated records with their source workbook and worksheet.
7. **Generate the output:** Write the resulting data to a new Excel workbook.

### Merge Workflow

```text
┌──────────────────────┐
│  Input Excel Files   │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Select Worksheets    │
│ and Matching Keys    │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Normalize Headers    │
│ and Key Values       │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Collect and Merge    │
│ Spreadsheet Records  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Apply Source Tracking│
│ and Cell Formatting  │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│ Consolidated Excel   │
│ Workbook             │
└──────────────────────┘
```

## Technology Stack

| Technology                           | Purpose                                                                         |
| ------------------------------------ | ------------------------------------------------------------------------------- |
| Python                               | Core application logic and data processing                                      |
| PySide6                              | Desktop graphical user interface                                                |
| openpyxl                             | Reading and writing Excel workbooks                                             |
| Python standard library              | File handling, regular expressions, resource management, and utility operations |
| PyInstaller-compatible specification | Application packaging configuration                                             |

## Project Structure

```text
excelmerger/
├── build/
│   └── MergeExcel/       # Packaged build artifacts
├── .gitignore            # Git ignore rules
├── LICENSE               # Apache License 2.0
├── MergeExcel.spec       # Application packaging specification
├── README.md             # Project documentation
├── icon.png              # Application icon
└── main.py               # Main application and merge logic
```

### Main Files

**`main.py`**

Contains the desktop interface, spreadsheet processing functions, column and key normalization, record consolidation, source tracking, and Excel output generation.

**`MergeExcel.spec`**

Defines the packaging configuration used to build a distributable application with PyInstaller.

**`icon.png`**

Contains the application icon resource.

**`build/`**

Contains packaged build artifacts. The exact contents depend on the build process.

## Prerequisites

Before running the application from source, make sure you have:

* Python installed.
* `pip` available for installing Python packages.
* A supported desktop operating system with the dependencies required by PySide6.
* Access to the Excel workbooks you want to consolidate.

Python 3.10 or later is a reasonable starting point for development, although the repository does not currently specify an explicit minimum Python version.

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/A-fawzy05/excelmerger.git
cd excelmerger
```

### 2. Create a virtual environment

Using a virtual environment isolates project dependencies from other Python projects.

**Windows — Command Prompt**

```bat
python -m venv .venv
.venv\Scripts\activate
```

**Windows — PowerShell**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**Linux or macOS**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

If PowerShell prevents activation because of its execution policy, consult the official Python and PowerShell documentation for the appropriate environment-specific solution.

### 3. Install dependencies

```bash
python -m pip install --upgrade pip
python -m pip install PySide6 openpyxl
```

The application imports PySide6 for its GUI and openpyxl for Excel file processing.

### 4. Launch the application

```bash
python main.py
```

The graphical application should open if the dependencies are installed correctly and your environment supports the GUI.

## Running the Application

### Running from source

For development or debugging, launch the application directly:

```bash
python main.py
```

This is the recommended approach when modifying the source code.

### Running a packaged build

The repository includes a PyInstaller specification and a build directory. If a runnable executable is available in the packaged artifacts, it can be launched directly on a compatible system.

Packaged executables are platform-dependent. An executable built for Windows should not be assumed to work on Linux or macOS.

## Using the Application

The exact appearance of the interface may vary between builds, but the intended workflow is based on selecting source files, configuring worksheet and matching options, and generating the consolidated workbook.

### Step 1: Prepare your Excel files

Place the workbooks you want to consolidate in an accessible directory.

For the most predictable results:

* Keep the relevant column headers in the first row of each worksheet.
* Use consistent identifiers wherever possible.
* Review duplicate or ambiguous headers.
* Keep a backup of important source files.

The current implementation reads `.xlsx` workbooks through openpyxl. Legacy `.xls` files are not supported by this implementation as written.

### Step 2: Select source workbooks and worksheets

Choose the Excel workbooks to include and select the worksheets to process.

The application uses worksheet names when deciding which sheets to include. Worksheets that are not selected are excluded from the consolidation.

### Step 3: Configure matching keys

If records should be combined based on a shared identifier, select the relevant key column or columns.

Examples include:

* Employee ID.
* Customer ID.
* Product code.
* Phone number.
* Another stable identifier shared across the source workbooks.

If multiple columns are selected, the combination of their values is used as the matching key.

Choose keys carefully. A key that is not sufficiently unique can cause unrelated records to be grouped together.

### Step 4: Run the merge

Start the merge operation and allow the application to process the selected worksheets.

The application aligns columns, groups matching records when keys are configured, and combines values according to its conflict-resolution rules.

### Step 5: Review the output

Inspect the resulting workbook before using it in another system or sharing it with others.

Pay particular attention to:

* The number of consolidated records.
* Missing values.
* Records with conflicting source data.
* The source-tracking column.
* Column headers and formatting.

## Merge Logic and Data Handling

Understanding the merge rules is important when working with business-critical data.

### Column alignment

The application builds a consolidated set of columns from the selected worksheets.

Columns are normalized before they are compared, allowing configured aliases and certain spelling variations to map to the same canonical column name.

Columns are placed in the output according to the order in which they are first encountered, with configured key columns moved to the beginning.

### Matching records

When key columns are configured, their values are normalized and combined to form a matching key.

For example:

| Source | Customer ID | Name      | Phone         |
| ------ | ----------- | --------- | ------------- |
| File A | `123`       | Ahmed     | *(empty)*     |
| File B | `١٢٣`       | *(empty)* | `01012345678` |

Because the numeric key representations normalize to the same value, these records can be grouped together.

The resulting record can contain the name from File A and the phone number from File B.

This example illustrates the matching logic; actual results depend on the selected keys and source data.

### Handling empty values and conflicts

The application follows a first-non-empty-value-wins strategy for cells belonging to matching records.

| Existing value | Incoming value  | Result                     |
| -------------- | --------------- | -------------------------- |
| Empty          | Populated       | Incoming value is used     |
| Populated      | Empty           | Existing value is retained |
| Populated      | Same value      | Existing value is retained |
| Populated      | Different value | Existing value is retained |

This behavior helps complete records with missing information without silently replacing an earlier populated value.

**Important:** The merge does not automatically resolve conflicting values intelligently. If two files contain different non-empty values for the same field, the earlier value wins. Review important conflicts before relying on the consolidated workbook.

### Rows without matching keys

If no key columns are configured, records are not grouped through key-based deduplication.

Likewise, a record with a missing value in one of the selected key columns is treated as a separate record rather than being grouped under an incomplete key.

### Duplicate headers

When a worksheet contains duplicate normalized column names, the implementation retains the first occurrence of that column within the row-processing logic.

For best results, remove duplicate or ambiguous column headers before merging.

### Original files

The application reads the input workbooks and writes a new output workbook rather than intentionally updating the original workbooks in place.

Nevertheless, keep backups of important source data and verify the generated workbook before deleting or replacing any originals.

## Arabic Data Normalization

A key feature of Excel Merger is its handling of common Arabic-language spreadsheet variations.

### Header normalization

The application applies several transformations when normalizing headers:

* Removes Arabic diacritics (tashkeel).
* Removes tatweel characters.
* Normalizes selected Arabic letter variants.
* Standardizes whitespace.
* Applies configured aliases to recognized column names.

For example, variations such as `أحمد` and `احمد` in a header are normalized to the same basic spelling.

The alias configuration can also map different names for the same field to a canonical header.

### Numeric key normalization

The application translates Arabic-Indic and Persian digits into Western digits when normalizing matching keys.

Examples:

| Original value | Normalized value |
| -------------- | ---------------- |
| `123`          | `123`            |
| `١٢٣`          | `123`            |
| `۱۲۳`          | `123`            |
| `123.0`        | `123`            |

This helps prevent otherwise equivalent identifiers from being treated as different records simply because their numeric representations differ.

Normalization is intentionally limited: it does not guarantee that every possible spelling, transliteration, or formatting difference will be recognized.

## Output Format

The generated workbook includes the consolidated columns and an additional source-tracking column.

The output worksheet is configured for right-to-left display.

The application also attempts to retain relevant source formatting, including header styling, cell borders, fills, alignments, number formats, and column widths.

### Important implementation details

* The source column is named `اسم_الملف`.
* Source labels include the workbook filename and worksheet name.
* The first encountered header supplies the formatting template for that column.
* The application reads workbook cell values using openpyxl's `data_only=True` mode.

**Formula limitation:** With `data_only=True`, formula cells are read using cached results saved in the workbook when available. If a workbook does not contain a usable cached result, a formula cell may not provide the calculated value you expect. This is especially important when consolidating formula-heavy spreadsheets.

The output should therefore be reviewed when source files contain formulas, complex workbook features, or data that requires exact preservation.

## Troubleshooting

### `ModuleNotFoundError: No module named 'PySide6'`

Install the GUI dependency in the active Python environment:

```bash
python -m pip install PySide6
```

### `ModuleNotFoundError: No module named 'openpyxl'`

Install the Excel-processing dependency:

```bash
python -m pip install openpyxl
```

### The application does not open

Verify that your virtual environment is activated and that the dependencies are installed.

Run the application from a terminal to inspect any error messages:

```bash
python main.py
```

If you are using a packaged executable, check whether the build matches your operating system and architecture.

### Records that should match remain separate

Check the following:

1. Confirm that the correct key columns are selected.
2. Verify that the key values are populated.
3. Check whether the values differ beyond the supported numeric normalization.
4. Inspect whitespace, punctuation, and other formatting differences.
5. Confirm that the relevant worksheets were selected.

The current key normalization is not a general-purpose fuzzy-matching algorithm.

### Different values appear in the consolidated workbook

The application retains the first populated value when matching records contain conflicting non-empty values.

Review the source workbooks and processing order to determine which value should be authoritative.

### Arabic headers are not aligned

Check whether the headers match the normalization rules or configured aliases in `main.py`.

You can extend the alias configuration to recognize additional known variations. Test the changes with representative input workbooks before relying on them in production.

### Formula results are missing or outdated

Open the original workbook in a spreadsheet application, recalculate it if appropriate, and save it before running the merge again.

Remember that the application relies on cached formula results rather than calculating Excel formulas itself.

### The output formatting differs from the original

The application copies selected cell styles and formatting properties. It does not guarantee complete preservation of every Excel workbook feature, such as macros, charts, external connections, or all advanced worksheet structures.

Use the generated workbook as a consolidated data output and validate any specialized formatting or workbook functionality that your workflow requires.

## Development and Contribution

Contributions, bug reports, and improvement suggestions are welcome.

### Getting started

1. Fork the repository.
2. Clone your fork.
3. Create and activate a virtual environment.
4. Install PySide6 and openpyxl.
5. Run the application with `python main.py`.
6. Make your changes and test them against representative Excel workbooks.
7. Submit a pull request describing the change and its expected behavior.

### Potential improvements

Possible future enhancements include:

* A dependency file such as `requirements.txt` for reproducible installation.
* Automated tests for header normalization and record matching.
* A preview of conflicting values before finalizing a merge.
* Configurable rules for resolving conflicting non-empty values.
* More configurable Arabic header aliases.
* Better reporting of skipped worksheets and invalid input files.
* Improved handling of formulas and unsupported workbook features.
* Packaging instructions for producing distributable executables.
* Additional validation for missing or non-unique matching keys.

These are suggestions for future development, not claims about existing functionality.

## License

This project is distributed under the **Apache License 2.0**. See the [`LICENSE`](LICENSE) file for the complete license terms.

## Author

**Amr Ahmed**

* GitHub: [@A-fawzy05](https://github.com/A-fawzy05)
* Repository: [excelmerger](https://github.com/A-fawzy05/excelmerger)

---

If you find a bug or have an idea for improving Excel Merger, consider opening an issue or contributing a change.

**Excel Merger — making spreadsheet consolidation simpler, more consistent, and easier to maintain.**
