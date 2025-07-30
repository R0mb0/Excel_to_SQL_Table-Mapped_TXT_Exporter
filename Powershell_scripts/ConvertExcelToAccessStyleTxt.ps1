<#
.SYNOPSIS
    Converts all XLSX files in the script directory to pipe-delimited .txt files (Access-style),
    according to column specifications/order read from a CSV file (typically produced by Export-SqlTableColumns_Version3.ps1).

.DESCRIPTION
    - All .xlsx files and the column properties file (.csv) must be in the same folder as the script.
    - For each Excel file, a .txt file is generated in the output subfolder, with no header,
      and with columns in the order and with the names defined in the CSV.
    - Values are pipe-delimited (|).
.REQUIREMENTS
    - ImportExcel PowerShell module installed.
#>

Write-Host "=== STARTING EXCEL → TXT (Access-style) CONVERSION ==="
# === Optional parameter: specify a worksheet name, or press Enter for the first sheet ===
$WorksheetName = Read-Host "Excel worksheet name to process (press Enter for first sheet)"

# === Path definitions ===
$ScriptDir = $PSScriptRoot
$OutputDir = Join-Path $ScriptDir "output"
$OutputDelimiter = "|"   # FIXED delimiter for Access-style

# Create output folder if it does not exist
if (-not (Test-Path $OutputDir)) {
    Write-Host "Creating output folder: $OutputDir"
    New-Item -Path $OutputDir -ItemType Directory | Out-Null
}

# Find the CSV file with column properties in the script folder
Write-Host "Looking for column property CSV file in folder: $ScriptDir"
$PropertiesCsv = Get-ChildItem -Path $ScriptDir -Filter *.csv | Select-Object -First 1
if (-not $PropertiesCsv) {
    Write-Error "No column property CSV file found in folder: $ScriptDir"
    exit 1
}
Write-Host "Column property file found: $($PropertiesCsv.Name)"

# Load column configuration (FIXED DELIMITER ; as in export script)
Write-Host "Loading column configuration from CSV..."
$ColProps = Import-Csv -Path $PropertiesCsv.FullName -Delimiter ';'
$ColumnOrder = $ColProps | Select-Object -ExpandProperty NomeColonna
Write-Host "Expected columns (order): $($ColumnOrder -join ', ')"

# Check ImportExcel
Write-Host "Checking for ImportExcel module..."
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "The ImportExcel module is not installed. Install it via: Install-Module -Name ImportExcel -Scope CurrentUser"
    exit 1
}
Import-Module ImportExcel -ErrorAction Stop

# Process all Excel files in the folder
Write-Host "Looking for .xlsx files in folder: $ScriptDir"
$ExcelFiles = Get-ChildItem -Path $ScriptDir -Filter *.xlsx
if (-not $ExcelFiles) {
    Write-Host "No .xlsx files found in $ScriptDir"
    exit 0
}
Write-Host "$($ExcelFiles.Count) files found to process."

foreach ($ExcelFile in $ExcelFiles) {
    Write-Host "---------------------------------------------"
    Write-Host "Processing Excel file: $($ExcelFile.Name)"

    # Import the specified worksheet, or first sheet if not specified
    try {
        if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
            Write-Host "Importing the first worksheet..."
            $Data = Import-Excel -Path $ExcelFile.FullName
        } else {
            Write-Host "Importing worksheet: $WorksheetName ..."
            $Data = Import-Excel -Path $ExcelFile.FullName -WorksheetName $WorksheetName
        }
    } catch {
        Write-Warning "Error reading $($ExcelFile.Name): $($_.Exception.Message)"
        continue
    }

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Warning "$($ExcelFile.Name) is empty or unreadable."
        continue
    }

    # Determine column names in the Excel file
    $ExcelColumns = $Data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    Write-Host "Columns found in file: $($ExcelColumns -join ', ')"

    # Check that all required columns are present; warn about missing columns but continue processing
    $MissingCols = @()
    foreach ($col in $ColumnOrder) {
        if (-not ($ExcelColumns -contains $col)) {
            $MissingCols += $col
        }
    }
    if ($MissingCols.Count -gt 0) {
        Write-Warning "Missing columns in $($ExcelFile.Name): $($MissingCols -join ', ') - Corresponding fields will be empty in the TXT."
    } else {
        Write-Host "All required columns are present."
    }

    # Write output .txt
    $TxtFileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelFile.Name) + ".txt"
    $TxtFilePath = Join-Path $OutputDir $TxtFileName

    Write-Host "Generating TXT file: $TxtFileName"
    $TxtContent = @()
    $rowNum = 0
    foreach ($row in $Data) {
        $fields = @()
        foreach ($col in $ColumnOrder) {
            if ($ExcelColumns -contains $col) {
                $raw = $row.$col
                if ($null -eq $raw) { $raw = "" }
                $fields += $raw
            } else {
                # Missing column: add empty field
                $fields += ""
            }
        }
        $TxtContent += ($fields -join $OutputDelimiter)
        $rowNum++
        if ($rowNum % 100 -eq 0) {
            Write-Host ("  ... rows processed: {0}" -f $rowNum)
        }
    }
    Write-Host ("Total rows processed: {0}" -f $rowNum)

    Set-Content -Path $TxtFilePath -Value $TxtContent -Encoding UTF8
    Write-Host "Created: $TxtFilePath"
}

Write-Host "=== CONVERSION COMPLETE. ==="