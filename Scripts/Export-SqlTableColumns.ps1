# Connection parameters - will be prompted to the user
Write-Host "=== EXPORT SQL SERVER TABLE COLUMNS PROPERTIES ==="

$SqlServer = Read-Host "Enter SQL Server name (e.g.: localhost\SQLEXPRESS)"
if ([string]::IsNullOrWhiteSpace($SqlServer)) {
    Write-Error "SQL Server name is required."
    exit 1
}

$Database = Read-Host "Enter database name"
if ([string]::IsNullOrWhiteSpace($Database)) {
    Write-Error "Database name is required."
    exit 1
}

$User = Read-Host "Enter SQL username"
if ([string]::IsNullOrWhiteSpace($User)) {
    Write-Error "SQL username is required."
    exit 1
}

$Password = Read-Host "Enter SQL password" -AsSecureString
if (-not $Password) {
    Write-Error "Password is required."
    exit 1
}
$PasswordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))

$TableName = Read-Host "Enter table name (case sensitive as in DB)"
if ([string]::IsNullOrWhiteSpace($TableName)) {
    Write-Error "Table name is required."
    exit 1
}

# Output directory (same as script location)
$OutputDir  = $PSScriptRoot
$OutputFile = Join-Path $OutputDir "table_columns_properties.csv"

# SQL query to get column properties
$Query = @"
SELECT
    COLUMN_NAME,
    DATA_TYPE,
    COALESCE(CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION, DATETIME_PRECISION, 0) AS COLUMN_LENGTH,
    IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = '$TableName'
ORDER BY ORDINAL_POSITION;
"@

# SQL Server connection string (SQL authentication)
$ConnectionString = "Server=$SqlServer;Database=$Database;User ID=$User;Password=$PasswordPlain;Trusted_Connection=False;"

# Connect and execute
$Connection = New-Object System.Data.SqlClient.SqlConnection
$Connection.ConnectionString = $ConnectionString

try {
    $Connection.Open()
} catch {
    Write-Error "Failed to connect to SQL Server. Please check your parameters. $($_.Exception.Message)"
    exit 1
}

$Command = $Connection.CreateCommand()
$Command.CommandText = $Query

try {
    $Reader = $Command.ExecuteReader()
} catch {
    Write-Error "Error executing the query. $($_.Exception.Message)"
    $Connection.Close()
    exit 1
}

# Write CSV header
"ColumnName;Type;Length;Nullable" | Out-File -Encoding UTF8 $OutputFile

# Loop through columns and write rows
while ($Reader.Read()) {
    $ColName = $Reader["COLUMN_NAME"]
    $ColType = $Reader["DATA_TYPE"]
    $ColLen  = $Reader["COLUMN_LENGTH"]
    $IsNull  = $Reader["IS_NULLABLE"]
    "$ColName;$ColType;$ColLen;$IsNull" | Out-File -Encoding UTF8 -Append $OutputFile
}

# Close everything
$Reader.Close()
$Connection.Close()

Write-Host "Done! Columns properties exported to $OutputFile"