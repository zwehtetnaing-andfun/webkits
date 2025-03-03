## Installation Require Dependencies

**Install `openpyxl`**

```sh
pip install openpyxl
```

**Install `pywin32`**

```sh
pip install pywin32
```

## Installation `.exe` file

**Install `pyinstaller`**

```sh
pip install pyinstaller
```

**Creation `.exe` file**

```sh
pyinstaller --onefile --noconsole --icon=your_icon.ico your_script.py
```

----
# Excel File Comparison Tool Documentation

## Overview

This program compares two Excel files containing work schedule data (EMEN), highlighting differences between them. It is tailored for Japanese workplace scheduling formats, with special attention to time formats, date handling, and workplace terminology.

## Core Features

### 1. Excel File Comparison
- Compares Excel files from two different directories
- Supports both `.xlsx` and `.xls` formats
- Preserves original file structure while marking differences
- Generates new files with differences highlighted in yellow

### 2. Special Handling
- Time format normalization (e.g., "08:00:00" → "8:00")
- Date comparison with multiple format support
- Time range comparison (using a support)
- Special text equivalence (e.g., "8:3p" = "5/7 1-8/10:2 - N-1")

### 3. Output Generation
- Creates files with `O_` prefix for matches
- Creates files with `X_` prefix for differences
- Highlights differences in yellow in the output file

## Technical Architecture

### Key Components

#### 1. File Selection System
- Uses `tkinter` for GUI file selection
- Implements a three-folder selection process:
  - Source folder 1 (VB1)
  - Source folder 2 (VB2)
  - Output destination folder (VB3)

#### 2. Logging System
- Configurable debug levels (`DEBUG`, `INFO`, `WARNING`)
- Timestamp-based log files
- Outputs to both file and console

#### 3. Excel Processing
- Uses `openpyxl` for Excel file handling
- Sheet comparison with name matching
- Cell-by-cell comparison with special rules

### Critical Functions

#### `compare_excel_files(file1_path, file2_path)`
- Main comparison function that:
  - Loads workbooks
  - Matches sheet names
  - Performs cell-by-cell comparison
  - Returns comparison result and modified workbook

#### `normalize_value(value)`
- Value normalization function handling:
  - `None` values
  - Datetime objects
  - Numeric values
  - Special characters
  - Whitespace

#### `get_comparison_columns(col, sheet_name, row)`
- Manages column mapping between sheets:
  - Handles special case for column 13
  - Adjusts column numbers for comparison
  - Returns tuple of `(col1, col2)` or `None` for skipped columns

## Extension Points

### Adding New Functionality

1. **New Value Type Handling**
   Add to `normalize_value()`:


# Excel File Comparison Tool

A Python-based tool designed to compare two Excel files containing work schedule data (EMEN), tailored for Japanese workplace scheduling formats. It highlights differences with special handling for time formats, dates, and workplace terminology.

---

## Table of Contents
- [Overview](#overview)
- [Core Features](#core-features)
- [Technical Architecture](#technical-architecture)
- [Installation](#installation)
- [Usage](#usage)
- [Extension Points](#extension-points)
- [Known Limitations](#known-limitations)
- [Maintenance Guidelines](#maintenance-guidelines)
- [Future Enhancements](#future-enhancements)
- [Dependencies](#dependencies)
- [Configuration](#configuration)
- [Modification Guide](#modification-guide)
- [Contributing](#contributing)

---

## Overview

This tool compares two Excel files, identifies differences, and generates highlighted output files. It’s optimized for Japanese workplace schedules, handling unique formatting like time normalization and special text equivalence.

## Core Features

### 1. Excel File Comparison
- Compares files from two directories
- Supports `.xlsx` and `.xls` formats
- Preserves original structure, highlighting differences in yellow
- Outputs new files with prefixes (`O_` for matches, `X_` for differences)

### 2. Special Handling
- Normalizes time formats (e.g., "08:00:00" → "8:00")
- Supports multiple date formats
- Compares time ranges
- Recognizes equivalent text (e.g., "8:3p" = "5/7 1-8/10:2 - N-1")

### 3. Output Generation
- Generates files with highlighted differences
- Uses yellow for mismatches in the output

## Technical Architecture

### Key Components
1. **File Selection System**
   - GUI via `tkinter`
   - Three-folder process: Source 1 (VB1), Source 2 (VB2), Output (VB3)

2. **Logging System**
   - Levels: `DEBUG`, `INFO`, `WARNING`
   - Timestamped logs to file and console

3. **Excel Processing**
   - Uses `openpyxl`
   - Matches sheets by name
   - Cell-by-cell comparison with custom rules

### Critical Functions
- **`compare_excel_files(file1_path, file2_path)`**: Loads workbooks, compares sheets, and highlights differences.
- **`normalize_value(value)`**: Normalizes values (e.g., `None`, datetimes, special characters).
- **`get_comparison_columns(col, sheet_name, row)`**: Maps columns, skips specific ones (e.g., column 13).

---

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/username/repo-name.git
   cd repo-name
   




