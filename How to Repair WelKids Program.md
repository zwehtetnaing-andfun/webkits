# Excel File Comparison Tool Documentation

## Overview
This program is designed to compare two Excel files containing work schedule data (勤務表), highlighting differences between them. It specifically handles Japanese
workplace scheduling formats with special attention to time formats, date handling, and specific workplace terminology.

## Core Feature

1. Excel File Comparison
    - Compares Excel files from two different directories
    - Supports both `.xlsx` and `.xls` formats
    - Preserves original file structure while marking differences
    - Generates new files with differences highlighted in yellow
2. Special Handling
    - Time format normalization (e.g., "08:00:00" → "8:00")
    - Date comparison with multiple format support
    - Time range comparison (using ~, 〜, ～ as separators)
    - Special text equivalence (e.g., "休み" = "シフト時間コード-1")
3. Output Generation
    - Creates files with 'O_' prefix for matches
    - Creates files with 'X_' prefix for differences
    - Highlights differences in yellow in the output file

## Technical Architecture
 

### Key Components

1. File Selection System
    - Uses `tkinter` for GUI file selection
    - Implements three-folder selection process:
         - Source folder 1 (VB1)
         - Source folder 2 (VB2)
         - Output destination folder (VB3)
2. Logging System
    - Configurable debug levels (DEBUG, INFO, WARNING)
    - Timestamp-based log files
    - Both file and console output
3. Excel ProcessingExcel Processing
    - Uses openpyxl for Excel file handling
    - Sheet comparison with name matching
    - Cell-by-cell comparison with special rules

### Critical Functions

```python
compare_excel_files(file1_path, file2_path)compare_excel_files(file1_path, file2_path)
```
Main comparison function that:
- Loads workbooks
- Matches sheet names
- Performs cell-by-cell comparison
- Returns comparison result and modified workbook

```python
normalize_value(value)normalize_value(value)
```
Value normalization function handling:
- None values
- Datetime objects
- Numeric values
- Special characters
- Whitespace

```python
get_comparison_columns(col, sheet_name, row)get_comparison_columns(col, sheet_name, row)
```
Manages column mapping between sheets:
- Handles special case for column 13
- Adjusts column numbers for comparison
- Returns tuple of (col1, col2) or None for skipped columns

## Extension Points
 

### Adding New Functionality

1. **New Value Type Handling** Add to `normalize_value()` function:

```python
def normalize_value(value):
    # Add new type handling here
    if isinstance(value, new_type):
        return processed_value
```

2. **New Comparison Rules** Extend `compare_excel_files()` function:

```python
# Add new comparison logic
if special_condition:
   # Handle special comparison
   pass
```

3. **Additional Output Formats** Modify the `main` function:

```python
# Add new output format handling
if special_format_needed:
   # Generate special format output
   pass
```
### Error Handling Expansion

Current error handling can be enhanced by:

1. Adding specific exception types
2. Implementing retry mechanisms
3. Adding transaction-like operations for file operations

## Known Limitations

 


1. Sheet Naming
   - Relies on sheet names for matching
   - Sensitive to leading numbers and symbols
2. Performance
   - Cell-by-cell comparison can be slow for large files
   - No parallel processing implementation
3. Memory Usage
   - Loads entire workbooks into memory
   - May struggle with very large Excel files

## Maintenance Guidelines
 

### Adding New Features


1. Always maintain logging:

```python
logging.debug('New feature debug info')
logging.info('New feature status')
logging.error('New feature error', exc_info=True)
```

2. Follow error handling pattern:
```python
try:
   # New feature code
except SpecificException as e:
   logging.error(f'Specific error: {str(e)}')
   show_message("Error", str(e))
except Exception as e:
   logging.error(f'Unexpected error: {str(e)}', exc_info=True)
   raise
```

### Testing New Features

1. Test with various Excel formats
2. Verify handling of:
   - Different time formats
   - Special characters
   - Empty cells
   - Large files

## Future Enhancement Suggestions
 

1. Performance Improvements
   - Implement parallel processing
   - Add batch processing capability
   - Optimize memory usage
2. User Interface Enhancements
   - Add progress bar
   - Implement cancel operation
   - Add detailed error reporting
3. Feature Additions
   - Add configuration file support
   - Implement custom comparison rules
   - Add report generation capability

## Dependencies



   - `Python 3.x`
   - `openpyxl`
   - `win32com.client`
   - `tkinter`
   - `logging`
   - `os`
   - `sys`
   - `re`
   - `datetime`

## Configuration
 

Logging levels can be set to:
- DEBUG: Detailed debugging information
- INFO: General operational information
- WARNING: Warning messages only

Example configuration change:
```python
setup_logging('DEBUG') # For detailed logging
setup_logging('INFO') # For normal operation
setup_logging('WARNING') # For minimal logging
```


## Modification Guide
 

### 1. Skipping Specific Columns

To skip specific columns during comparison, modify the `get_comparison_columns()` function:


```python
def get_comparison_columns(col, sheet_name, row):
   # Add new columns to skip
   columns_to_skip = [13, 25, 30] # Example: skip columns 13, 25, and 30
   if col in columns_to_skip:
      logging.debug(f'Column {col} is in skip list, skipping comparison')
      return None
# Rest of the existing function...
```
Location: Find the function` get_comparison_columns()` around line 240 in `kinmu.py`

### 2. Skipping Specific Rows

To skip rows, modify the comparison loop in `compare_excel_files()`:

```python
def compare_excel_files(file1_path, file2_path):
   # Existing code...

   # Add row skip conditions
   def should_skip_row(row_num, sheet1, sheet2):
      # Example: Skip header rows (first 3 rows)
      if row_num <= 3:
         return True

      # Example: Skip empty rows
      if not any(sheet1.cell(row_num, col).value for col in range(1, sheet1.max_column + 1)):
         return True

      # Example: Skip rows with specific values
      if sheet1.cell(row_num, 1).value == "小計":
         return True

      return False

   # In the main comparison loop
   for row in range(1, row_max + 1):
      if should_skip_row(row, sheet1, sheet2):
         logging.debug(f'Skipping row {row}')
         continue
```
Location: Find the function `compare_excel_files()` around line 390 in `kinmu.py`

### 3. Limiting Maximum Rows and Columns

To limit the maximum rows and columns for comparison, modify the dimension calculation in compare_excel_files():

```python
def compare_excel_files(file1_path, file2_path):
   # Existing code...

   # Add maximum limits
   MAX_ROWS = 100 # Example: limit to first 100 rows
   MAX_COLS = 30 # Example: limit to first 30 columns

   # Modify the row and column max calculations
   row_max = min(
      max(sheet1.max_row, sheet2.max_row),
      MAX_ROWS
   )
   col_max = min(
      max(sheet1.max_column, sheet2.max_column),
      MAX_COLS
   )
```

Location: Find the max row/column calculation in `compare_excel_files()` around line 450 in `kinmu.py`

### 4. Modifying Time Format Comparison

To change how time formats are compared, modify the normalize_time_format() function:

```python
def normalize_time_format(time_str):
   # Add new time format handling
   special_times = {
      "24:00": "0:00",
      "25:00": "1:00",
      # Add more special cases
   }

   if time_str in special_times:
      return special_times[time_str]
```
Location: Find the function `normalize_time_format()` around line 280 in `kinmu.py`

### 5. Adding New Value Comparison Rules

To add new rules for comparing values, modify the comparison section in `compare_excel_files()`:

```python
def compare_excel_files(file1_path, file2_path):
   # In the cell comparison loop

   # Add new comparison rules
   def custom_comparison(value1, value2):
      # Example: Treat specific values as equivalent
      equivalent_values = {
         "休暇": ["年休", "有休", "休暇"],
         "欠勤": ["欠勤", "欠"],
         # Add more equivalence groups
      }

      for group in equivalent_values.values():
         if value1 in group and value2 in group:
            return True
      return False

   # In the comparison logic
   if custom_comparison(value1, value2):
      continue
```
Location: Find the value comparison section in `compare_excel_files()` around line 500 in `kinmu.py`

### 6. Modifying Sheet Name Matching

To change how sheets are matched between workbooks, modify the sheet name extraction:

```python
def extract_sheet_name_string(sheet_name):
   # Add custom sheet name matching rules
   # Example: Remove specific prefixes
   prefixes_to_remove = ["Sheet", "シート", "Copy of"]

   result = sheet_name
   for prefix in prefixes_to_remove:
      if result.startswith(prefix):
         result = result[len(prefix):].strip()

   # Remove numbers and symbols as before
   result = re.sub(r'^[\d._\-]+', '', result)

   return result
```

Location: Find the function `extract_sheet_name_string()` around line 350 in `kinmu.py`

### 7. Adding New Ignored Mismatches

To add new pairs of values that should be considered equivalent:

```python
def is_ignored_mismatch(value1, value2):
   ignored_pairs = [
      ("休み", "シフト時間コード-1"),
      ("フリー", "シフト時間コード2147483647"),
      # Add new pairs here:
      ("出張", "外出"),
      ("在宅", "リモート"),
      # More pairs...
   ]
   return (value1, value2) in ignored_pairs or (value2, value1) in ignored_pairs
```
Location: Find the function `is_ignored_mismatch()` around line 580 in `kinmu.py`

### 8. Modifying Output Highlighting

To change how differences are highlighted in the output:

```python
def compare_excel_files(file1_path, file2_path):
   # Define different highlight colors
   highlight_colors = {
      'mismatch': PatternFill(patternType="solid", fgColor='FFFF00'), # Yellow
      'missing': PatternFill(patternType="solid", fgColor='FF0000'), # Red
      'added': PatternFill(patternType="solid", fgColor='00FF00'), # Green
   }

   # In the comparison logic
   if value1 is None and value2 is not None:
      sheet2.cell(row, col2).fill = highlight_colors['added']
   elif value1 is not None and value2 is None:
      sheet2.cell(row, col2).fill = highlight_colors['missing']
   elif value1 != value2:
      sheet2.cell(row, col2).fill = highlight_colors['mismatch']
```
Location: Find the highlighting code in `compare_excel_files()` around line 460 in `kinmu.py`

Remember to always:


1. Add appropriate logging for new changes
2. Test thoroughly with different Excel files
3. Update the documentation when adding new features
4. Consider backwards compatibility when making changes

Would you like me to elaborate on any of these modification points or provide additional examples for other common scenarios?


