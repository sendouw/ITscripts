# Simple Database Query Test
# This will show us exactly what's happening with the query

param(
    [string]$DSN = "LAB_PRIMARY"
)

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Database Query Diagnostic Test" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

try {
    Write-Host "Step 1: Connecting to DSN: $DSN" -ForegroundColor Yellow
    $connection = New-Object System.Data.Odbc.OdbcConnection("DSN=$DSN;Trusted_Connection=Yes;")
    $connection.Open()
    Write-Host "  SUCCESS: Connected!" -ForegroundColor Green
    Write-Host "  Database: $($connection.Database)" -ForegroundColor White
    Write-Host ""
    
    # Test 1: Simple count query
    Write-Host "Step 2: Testing simple COUNT query on dbo.tb_Technician" -ForegroundColor Yellow
    $countQuery = "SELECT COUNT(*) as TotalRecords FROM dbo.tb_Technician"
    $command = $connection.CreateCommand()
    $command.CommandText = $countQuery
    
    try {
        $reader = $command.ExecuteReader()
        if ($reader.Read()) {
            $count = $reader["TotalRecords"]
            Write-Host "  SUCCESS: Found $count records" -ForegroundColor Green
        }
        $reader.Close()
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $connection.Close()
        Read-Host "Press Enter to exit"
        exit
    }
    Write-Host ""
    
    # Test 2: Get first 5 records
    Write-Host "Step 3: Querying first 5 technicians" -ForegroundColor Yellow
    $query = @"
SELECT TOP 5
    T.TechID,
    T.UserName,
    T.Area,
    T.Name
FROM dbo.tb_Technician T
ORDER BY T.UserName
"@
    
    $command2 = $connection.CreateCommand()
    $command2.CommandText = $query
    
    $adapter = New-Object System.Data.Odbc.OdbcDataAdapter($command2)
    $results = New-Object System.Data.DataTable
    
    Write-Host "  Executing query..." -ForegroundColor White
    $rowCount = $adapter.Fill($results)
    Write-Host "  SUCCESS: Adapter.Fill returned $rowCount rows" -ForegroundColor Green
    Write-Host "  Results.Rows.Count: $($results.Rows.Count)" -ForegroundColor White
    Write-Host ""
    
    if ($results.Rows.Count -gt 0) {
        Write-Host "Step 4: Displaying results" -ForegroundColor Yellow
        Write-Host "  Columns: $($results.Columns.Count)" -ForegroundColor White
        foreach ($col in $results.Columns) {
            Write-Host "    - $($col.ColumnName) ($($col.DataType.Name))" -ForegroundColor Cyan
        }
        Write-Host ""
        
        Write-Host "  First record:" -ForegroundColor White
        $firstRow = $results.Rows[0]
        Write-Host "    TechID: $($firstRow['TechID'])" -ForegroundColor Cyan
        Write-Host "    UserName: $($firstRow['UserName'])" -ForegroundColor Cyan
        Write-Host "    Area: $($firstRow['Area'])" -ForegroundColor Cyan
        Write-Host "    Name: $($firstRow['Name'])" -ForegroundColor Cyan
    }
    else {
        Write-Host "  WARNING: No rows returned!" -ForegroundColor Red
    }
    Write-Host ""
    
    # Test 3: Try the JOIN query
    Write-Host "Step 5: Testing JOIN query with TechVerify" -ForegroundColor Yellow
    $joinQuery = @"
SELECT TOP 5
    T.TechID,
    T.UserName,
    T.Area,
    T.Name,
    V.FirstName,
    V.LastName
FROM dbo.tb_Technician T
LEFT JOIN dbo.tb_TechVerify V ON T.TechID = V.TechID
ORDER BY T.UserName
"@
    
    $command3 = $connection.CreateCommand()
    $command3.CommandText = $joinQuery
    
    $adapter2 = New-Object System.Data.Odbc.OdbcDataAdapter($command3)
    $results2 = New-Object System.Data.DataTable
    
    $rowCount2 = $adapter2.Fill($results2)
    Write-Host "  SUCCESS: JOIN query returned $rowCount2 rows" -ForegroundColor Green
    
    if ($results2.Rows.Count -gt 0) {
        $firstRow2 = $results2.Rows[0]
        Write-Host "  First record:" -ForegroundColor White
        Write-Host "    TechID: $($firstRow2['TechID'])" -ForegroundColor Cyan
        Write-Host "    UserName: $($firstRow2['UserName'])" -ForegroundColor Cyan
        Write-Host "    FirstName: $($firstRow2['FirstName'])" -ForegroundColor Cyan
        Write-Host "    LastName: $($firstRow2['LastName'])" -ForegroundColor Cyan
    }
    Write-Host ""
    
    # Test 4: Check privileges table
    Write-Host "Step 6: Testing TechPrivilege table" -ForegroundColor Yellow
    $privQuery = "SELECT TOP 5 TechID, Privilege FROM dbo.tb_TechPrivilege"
    
    $command4 = $connection.CreateCommand()
    $command4.CommandText = $privQuery
    
    $adapter3 = New-Object System.Data.Odbc.OdbcDataAdapter($command4)
    $results3 = New-Object System.Data.DataTable
    
    $rowCount3 = $adapter3.Fill($results3)
    Write-Host "  SUCCESS: Found $rowCount3 privilege records" -ForegroundColor Green
    
    if ($results3.Rows.Count -gt 0) {
        $firstPriv = $results3.Rows[0]
        Write-Host "  First privilege record:" -ForegroundColor White
        Write-Host "    TechID: $($firstPriv['TechID'])" -ForegroundColor Cyan
        Write-Host "    Privilege: $($firstPriv['Privilege'])" -ForegroundColor Cyan
    }
    Write-Host ""
    
    $connection.Close()
    
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host "All tests PASSED!" -ForegroundColor Green
    Write-Host "=====================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "The database queries work fine." -ForegroundColor Yellow
    Write-Host "The error is likely in how the AccessManager script" -ForegroundColor Yellow
    Write-Host "is processing the results." -ForegroundColor Yellow
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Red
    Write-Host "ERROR OCCURRED" -ForegroundColor Red
    Write-Host "=====================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error Message:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor White
    Write-Host ""
    Write-Host "Error Type:" -ForegroundColor Red
    Write-Host $_.Exception.GetType().FullName -ForegroundColor White
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor White
}

Write-Host ""
Read-Host "Press Enter to exit"
