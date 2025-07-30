# Excel to SQL Table-Mapped TXT Exporter

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/92f0f1cd0a6749738580072486da0f32)](https://app.codacy.com/gh/R0mb0/Excel_to_SQL_Table-Mapped_TXT_Exporter/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Excel_to_SQL_Table-Mapped_TXT_Exporter)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Excel_to_SQL_Table-Mapped_TXT_Exporter)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

Convert Excel (.xlsx) files into pipe-delimited TXT files, automatically mapping columns and order to the structure of an existing SQL Server table.  
This utility is designed for seamless integration with legacy systems, automated data imports, and ETL pipelines—especially when you need to generate TXT files compatible with classic ASP, Access, or other systems with fixed text-import requirements.

## Overview

Many organizations receive data in Excel format that must be imported into databases or legacy systems requiring a specific TXT structure.  
If you need to ensure that your exported TXT files match the schema of a SQL Server table (column names, order, and types), this project provides a robust, script-based solution that automates the mapping and conversion process.

## Key Features

- **Automatic SQL Schema Detection**: Reads column names, types, and order from an existing SQL Server table and exports them to a CSV file.
- **Excel to TXT Conversion**: Converts one or more Excel files to pipe-delimited TXT files, matching the required schema exactly.
- **Batch Processing**: Handles multiple Excel files in a folder automatically.
- **Legacy Compatibility**: Output format is compatible with systems expecting "Access-style" TXT files (no header row, fixed column order, empty fields for missing columns).
- **Easy Adaptation**: Suitable for data migration, ETL, and legacy system integration.

## Typical Use Cases

- **Legacy System Integration**: When you must deliver TXT files for import by classic ASP, Access, or mainframe systems that expect a rigid text format.
- **Automated Data Imports**: For scenarios where you receive Excel files from third parties and must automate their import into a database.
- **ETL Pipelines**: As a preprocessing step to standardize data before further processing.
- **Format Normalization**: When you need to ensure Excel files from diverse sources conform to a master schema.

## Workflow

1. **Extract Table Structure from SQL Server**

    Use `Export-SqlTableColumns.ps1` to connect to your SQL Server and export the schema of the desired table (columns, types, order) to a CSV file.  
    This CSV acts as the "blueprint" for subsequent TXT exports.

2. **Convert Excel Files to TXT**

    Place your Excel files in the same folder as the scripts.  
    Use `ConvertExcelToAccessStyleTxt.ps1` to process each Excel file and generate a TXT file for each, formatted and ordered as required by the table schema from the CSV.

3. **Import TXT into Your System**

    The resulting TXT files can now be imported into your legacy system, classic ASP, Access, or any process that expects this format.

## Step-by-Step Usage

### 1. Requirements

- Windows with PowerShell 5.0 or later
- [ImportExcel PowerShell module](https://github.com/dfinke/ImportExcel) (install with `Install-Module -Name ImportExcel -Scope CurrentUser`)
- Access to the target SQL Server
- Set the appropriate PowerShell execution policy if needed (see below)

### 2. Extract SQL Table Schema

Run the following script and provide the requested SQL Server connection parameters and table name:

```powershell
.\Export-SqlTableColumns.ps1
```

This will generate a CSV file (e.g., `table_columns.csv`) containing the column definitions of your SQL table.

### 3. Prepare Excel Files

Copy all Excel files (.xlsx) you wish to convert into the same folder as the scripts and the CSV schema file.

### 4. Convert Excel to TXT

Run the conversion script:

```powershell
.\ConvertExcelToAccessStyleTxt.ps1
```

- The script will prompt for the worksheet name (press Enter for the first worksheet).
- TXT files will be generated in the `output/` subfolder, with one TXT per Excel file.
- Each TXT will have columns ordered and named according to the SQL schema, with missing fields left empty.

### 5. (Optional) Set Execution Policy

If you encounter permissions issues running PowerShell scripts, you can temporarily allow execution for the current process:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Example Folder Structure

```
YourFolder/
│   Export-SqlTableColumns.ps1
│   ConvertExcelToAccessStyleTxt.ps1
│   table_columns.csv
│   Example1.xlsx
│   Example2.xlsx
│
└───output/
        Example1.txt
        Example2.txt
```

## Script Details

### Export-SqlTableColumns.ps1

- Prompts the user for SQL Server connection info and target table name.
- Exports the table structure (column names, types, order) to CSV.
- CSV is used as the mapping template for TXT export.

### ConvertExcelToAccessStyleTxt.ps1

- Prompts for worksheet name, reads all .xlsx files in the folder.
- Uses the CSV to determine required columns and order.
- Produces pipe-delimited TXT files in the output folder, with values aligned to the schema.

## Troubleshooting

- Ensure the [ImportExcel](https://github.com/dfinke/ImportExcel) module is installed before running the conversion script.
- Column names in Excel must exactly match those in the SQL table (case-sensitive).
- Any missing columns in Excel will result in empty fields in the TXT file.
- If you receive "script not allowed to run" errors, use the execution policy command above.
