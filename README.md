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

#Excel File Comparison Tool Documentation

**Overview**

_This program is designed to compare two Excel files containing work schedule data (勤務表), highlighting differences between them. It specifically handles Japanese
workplace scheduling formats with special attention to time formats, date handling, and specific workplace terminology._

**Core Features**

1. Excel File Comparison
-Compares Excel files from two different directories
-Supports both `.xlsx` and `.xls` formats
-Preserves original file structure while marking differences
-Generates new files with differences highlighted in yellow
2. Special Handling
-Time format normalization (e.g., `"08:00:00` → `8:00`)
-Date comparison with multiple format support
-Time range comparison (using `~`, `〜`, `～` as separators)
-Special text equivalence (e.g., `休み` = `シフト時間コード-1`)
3. Output Generation
-Creates files with `O_` prefix for matches
-Creates files with `X_` prefix for differences
-Highlights differences in yellow in the output file




