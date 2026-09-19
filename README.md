# ExcelUtils

A legacy PHP utility for importing and generating XLS, XLSX and CSV files with PhpSpreadsheet.

I keep this repository public as a compact example of older PHP integration work, especially around spreadsheet generation, export formatting and compatibility with PHP 5.6-era systems.

## What it does

- Imports spreadsheet data into PHP arrays
- Generates XLS files
- Generates XLSX files
- Generates CSV files
- Supports configurable column mapping
- Supports simple value transformations and callbacks
- Supports result sets returned by SQL Server integrations

## Main file

```text
class_utils_excel.php
```

The class wraps PhpSpreadsheet operations and writes generated documents directly to the HTTP response.

## Historical context

This utility targets an older PHP environment and should be treated as legacy code. It is included in my public GitHub to document the kind of integration and maintenance work I have handled across different generations of PHP applications.
