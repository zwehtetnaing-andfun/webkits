**Exce**ll **F**iill**e** **Compar**ii**son** **Too**ll
**Documentat**ii**on**

> **Overv**ii**ew**
>
> This program is designed to compare two Excel files containing work
> schedule data (勤務表), highlighting differences between them. It
> specifically handles Japanese workplace scheduling formats with
> special attention to time formats, date handling, and specific
> workplace terminology.
>
> **Core** **Features**
>
> 1\. Excel File Comparison
>
> Compares Excel files from two different directories Supports both
> .xlsx and .xls formats
>
> Preserves original file structure while marking differences Generates
> new files with differences highlighted in yellow
>
> 2\. Special Handling
>
> Time format normalization (e.g., "08:00:00" → "8:00") Date comparison
> with multiple format support
>
> Time range comparison (using ~, 〜, ～ as separators) Special text
> equivalence (e.g., "休み" = "シフト時間コード-1")
>
> 3\. Output Generation
>
> Creates files with 'O\_' prefix for matches Creates files with 'X\_'
> prefix for differences Highlights differences in yellow in the output
> file
>
> **Techn**ii**ca**ll **Arch**ii**tecture**
>
> **Key** **Components**
>
> 1\. **F**iill**e** **Se**ll**ec**ttii**on** **Sys**tt**em**
>
> Uses tkinter for GUI file selection Implements three-folder selection
> process:
>
> Source folder 1 (VB1) Source folder 2 (VB2)
>
> Output destination folder (VB3)
>
> 2\. **Logg**ii**ng** **Sys**tt**em**
>
> Configurable debug levels (DEBUG, INFO, WARNING) Timestamp-based log
> files
>
> Both file and console output
>
> 3\. **Exce**ll **P**rr**ocess**ii**ng**
>
> Uses openpyxl for Excel file handling Sheet comparison with name
> matching Cell-by-cell comparison with special rules
>
> **Cr**ii**t**ii**ca**ll **Funct**ii**ons**
>
> **compare_excel_files(file1_path,** **file2_path)**
>
> Main comparison function that:
>
> Loads workbooks Matches sheet names
>
> Performs cell-by-cell comparison
>
> Returns comparison result and modified workbook
>
> **normalize_value(value)**
>
> Value normalization function handling:
>
> None values Datetime objects Numeric values Special characters
> Whitespace
>
> **get_comparison_columns(col,** **sheet_name,** **row)**
>
> Manages column mapping between sheets:
>
> Handles special case for column 13
>
> Adjusts column numbers for comparison
>
> Returns tuple of (col1, col2) or None for skipped columns

**Extens**ii**on** **Po**ii**nts**

**Add**ii**ng** **New** **Funct**ii**ona**llii**ty**

> 1\. **New** **Va**ll**ue** **Type** **Hand**llii**ng** Add to
> normalize_value() function:

||
||
||
||
||
||

> 2\. **New** **Compa**rrii**son** **Ru**ll**es** Extend
> compare_excel_files() function:

||
||
||
||
||
||

> 3\. **Add**iittii**ona**ll **Ou**tt**pu**tt **Fo**rr**ma**tt**s**
> Modify the main function:

||
||
||
||
||
||

**Error** **Hand**llii**ng** **Expans**ii**on**

> Current error handling can be enhanced by:
>
> 1\. Adding specific exception types 2. Implementing retry mechanisms
>
> 3\. Adding transaction-like operations for file operations

**Known** **L**ii**m**ii**tat**ii**ons**

> 1\. Sheet Naming
>
> Relies on sheet names for matching Sensitive to leading numbers and
> symbols
>
> 2\. Performance
>
> Cell-by-cell comparison can be slow for large files No parallel
> processing implementation
>
> 3\. Memory Usage
>
> Loads entire workbooks into memory May struggle with very large Excel
> files

**Ma**ii**ntenance** **Gu**ii**de**llii**nes**

**Add**ii**ng** **New** **Features**

> 1\. Always maintain logging:

||
||
||
||
||

> 2\. Follow error handling pattern:

||
||
||
||
||
||
||
||
||
||

**Test**ii**ng** **New** **Features**

> 1\. Test with various Excel formats 2. Verify handling of:
>
> Different time formats Special characters Empty cells
>
> Large files

**Future** **Enhancement** **Suggest**ii**ons**

> 1\. Performance Improvements
>
> Implement parallel processing Add batch processing capability Optimize
> memory usage
>
> 2\. User Interface Enhancements
>
> Add progress bar Implement cancel operation Add detailed error
> reporting
>
> 3\. Feature Additions
>
> Add configuration file support Implement custom comparison rules Add
> report generation capability

**Dependenc**ii**es**

> Python 3.x openpyxl win32com.client tkinter
>
> logging os
>
> sys re
>
> datetime

**Conf**ii**gurat**ii**on**

> Logging levels can be set to:
>
> DEBUG: Detailed debugging information INFO: General operational
> information WARNING: Warning messages only
>
> Example configuration change:

||
||
||
||

**Mod**ii**f**ii**cat**ii**on** **Gu**ii**de**

**1**.. **Sk**ii**pp**ii**ng** **Spec**ii**f**ii**c** **Co**ll**umns**

> To skip specific columns during comparison, modify the
> get_comparison_columns() function:

||
||
||
||
||
||
||
||
||
||
||

> Location: Find the function get_comparison_columns() around line 240
> in kinmu.py

**2**.. **Sk**ii**pp**ii**ng** **Spec**ii**f**ii**c** **Rows**

> To skip rows, modify the comparison loop in compare_excel_files():

||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||

> Location: Find the function compare_excel_files() around line 390 in
> kinmu.py

**3**.. **L**ii**m**ii**t**ii**ng** **Max**ii**mum** **Rows** **and**
**Co**ll**umns**

> To limit the maximum rows and columns for comparison, modify the
> dimension calculation in compare_excel_files():

||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||

> Location: Find the max row/column calculation in compare_excel_files()
> around line 450 in kinmu.py

**4**.. **Mod**ii**fy**ii**ng** **T**ii**me** **Format**
**Compar**ii**son**

> To change how time formats are compared, modify the
> normalize_time_format() function:

||
||
||
||
||
||
||
||
||
||
||

> Location: Find the function normalize_time_format() around line 280 in
> kinmu.py

**5**.. **Add**ii**ng** **New** **Va**ll**ue** **Compar**ii**son**
**Ru**ll**es**

> To add new rules for comparing values, modify the comparison section
> in compare_excel_files():

||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||

> Location: Find the value comparison section in compare_excel_files()
> around line 500 in kinmu.py

**6**.. **Mod**ii**fy**ii**ng** **Sheet** **Name** **Match**ii**ng**

> To change how sheets are matched between workbooks, modify the sheet
> name extraction:

||
||
||
||
||
||
||
||
||
||
||
||
||
||

> Location: Find the function extract_sheet_name_string() around line
> 350 in kinmu.py

**7**.. **Add**ii**ng** **New** II**gnored** **M**ii**smatches**

> To add new pairs of values that should be considered equivalent:

||
||
||
||
||
||
||
||
||
||
||
||

> Location: Find the function is_ignored_mismatch() around line 580 in
> kinmu.py

**8**.. **Mod**ii**fy**ii**ng** **Output**
**H**ii**gh**llii**ght**ii**ng**

> To change how differences are highlighted in the output:

||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||
||

> Location: Find the highlighting code in compare_excel_files() around
> line 460 in kinmu.py
>
> Remember to always:
>
> 1\. Add appropriate logging for new changes 2. Test thoroughly with
> different Excel files
>
> 3\. Update the documentation when adding new features
>
> 4\. Consider backwards compatibility when making changes
>
> Would you like me to elaborate on any of these modification points or
> provide additional examples for other common scenarios?
