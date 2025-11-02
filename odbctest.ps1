# --- SETTINGS (change if needed) ---
$dsn = "LAB_PRIMARY"   # Run again with LAB_SECONDARY for SECONDARY

# --- OPEN ODBC ---
Add-Type -AssemblyName System.Data
$c = New-Object System.Data.Odbc.OdbcConnection("DSN=$dsn;Trusted_Connection=Yes;")
$c.Open(); "Connected to DSN=$dsn"

# --- 1) LIST TABLES/VIEW NAMES (helps confirm real names) ---
$cmd = $c.CreateCommand()
$cmd.CommandText = @"
SELECT TOP 500 TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE
FROM INFORMATION_SCHEMA.TABLES
ORDER BY TABLE_SCHEMA, TABLE_NAME
"@
$adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
$dt  = New-Object System.Data.DataTable
[void]$adp.Fill($dt)
"=== TABLES (first 500) ==="
$dt | Format-Table -AutoSize

# --- 2) TRY TO GUESS CANDIDATES BY NAME ---
function Find-Candidates([string]$like){
  $cmd = $c.CreateCommand()
  $cmd.CommandText = @"
SELECT TABLE_SCHEMA, TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_NAME LIKE '$like'
ORDER BY TABLE_SCHEMA, TABLE_NAME
"@
  $adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
  $t   = New-Object System.Data.DataTable
  [void]$adp.Fill($t)
  return $t
}
"=== Candidates: *Technician ===";  Find-Candidates "%Technician%" | ft -AutoSize
"=== Candidates: *TechPrivilege ===";Find-Candidates "%TechPrivilege%" | ft -AutoSize
"=== Candidates: *TechVerify ===";   Find-Candidates "%TechVerify%"  | ft -AutoSize
"=== Candidates: *LabOps%Live% ===";  Find-Candidates "LabOps%Live%"  | ft -AutoSize

# --- 3) SCHEMA (COLUMNS) DUMPS for likely defaults (adjust after you see names) ---
$schema = "dbo"                # change after you see TABLE_SCHEMA
$tech   = "tb_Technician"      # change if your list shows a different name
$priv   = "tb_TechPrivilege"
$verify = "tb_TechVerify"
$mlive  = "tb_LabOpsLive"      # if it exists

function Show-Cols($schema, $table){
  $cmd = $c.CreateCommand()
  $cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA='$schema' AND TABLE_NAME='$table'
ORDER BY ORDINAL_POSITION
"@
  $adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
  $t   = New-Object System.Data.DataTable
  [void]$adp.Fill($t)
  "=== $schema.$table Columns ==="
  $t | ft -AutoSize
}

Show-Cols $schema $tech
Show-Cols $schema $priv
Show-Cols $schema $verify
Show-Cols $schema $mlive

# --- 4) PARAMETERIZED PROBE (checks the ? placeholders work) ---
# Adjust names if your tables differ:
$cmd = $c.CreateCommand()
$cmd.CommandText = @"
SELECT TOP 5 T.TechID, T.UserName
FROM [$schema].[$tech] T
WHERE (T.UserName LIKE ? OR CAST(T.TechID AS VARCHAR(20)) LIKE ?)
ORDER BY T.UserName
"@
$p1 = $cmd.CreateParameter(); $p1.Value = "%A%"; [void]$cmd.Parameters.Add($p1)
$p2 = $cmd.CreateParameter(); $p2.Value = "%1%"; [void]$cmd.Parameters.Add($p2)
$adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
$res = New-Object System.Data.DataTable
[void]$adp.Fill($res)
"=== Parameterized test (TOP 5) ==="
$res | ft -AutoSize

$c.Close()
