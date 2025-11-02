<#
.SYNOPSIS
    Access Manager - Production RBAC Tool for LABOPS/OpsPortal Access Databases
    
.DESCRIPTION
    A single-file PowerShell/WPF application for managing technician access in
    LABOPS and OpsPortal systems via Access database backend (PRIMARY/SECONDARY platforms).
    
    **FIXED:** Now supports ODBC/DSN connections for linked table databases
    
    Your DSNs: LAB_PRIMARY (PRIMARY) and LAB_SECONDARY (SECONDARY)
    
.NOTES
    Version: 1.1.1
    Last Updated: 2025-10-29
    
    Changes:
    - FIXED: ODBC result processing for WPF DataGrid binding
    - Improved error handling for ODBC connections
    - Better null value handling
    
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$GenerateSampleData,
    
    [Parameter()]
    [string]$DatabasePath
)

#Requires -Version 5.1

# ============================================================================
# GLOBAL CONFIGURATION AND CONSTANTS
# ============================================================================

$Script:AppVersion = "1.1.1"
$Script:AppName = "Access Manager"
$Script:ConfigPath = "$env:APPDATA\AccessManager\config.json"
$Script:DefaultConfig = @{
    DefaultDatabasePath = ""
    ConnectionMode = "ODBC"  # "File" or "ODBC"
    ODBCDSN_PRIMARY = "LAB_PRIMARY"
    ODBCDSN_SECONDARY = "LAB_SECONDARY"
    LoggingPath = "$env:USERPROFILE\Documents\AccessManager\Audit"
    BackupPath = ""
    BackupRetentionDays = 30
    AllowedPrivileges = @("VIEW", "ACCEPT", "APPROVE", "REAPPROVE", "INTERPRETATION", "LIMITA", "LIMITB")
    AllowedAreas = @("SPECHEM", "ENDO", "STEROIDS", "TOXICOLOGY", "IMMCHEM", "BIOCHEMG", "THYROID", "TIP", "TM", "LCCORE", "SEROLOGY", "INFORMATION", "MOLMICRO", "CLSID", "labops")
    AutoDetectConfigPath = "C:\LabOps\config.ini"
    LockTimeout = 30
    EnableAuditLog = $true
    EnableWindowsEventLog = $true
    MaxSearchResults = 1000
    PageSize = 50
    ReadOnlyMode = $false
}

$Script:Config = $null
$Script:CurrentConnection = $null
$Script:CurrentPlatform = $null
$Script:CurrentDatabasePath = $null
$Script:CurrentConnectionMode = "ODBC"
$Script:WhatIfMode = $false
$Script:LastOperation = $null
$Script:AppLock = $null
$Script:ReadOnlyMode = $false

# ODBC Drivers
$Script:ODBCDrivers = @(
    "ODBC Driver 17 for SQL Server",
    "ODBC Driver 13 for SQL Server",
    "SQL Server Native Client 11.0",
    "SQL Server"
)

# ============================================================================
# CONFIGURATION MANAGEMENT
# ============================================================================

function Initialize-Configuration {
    if (Test-Path $Script:ConfigPath) {
        try {
            $Script:Config = Get-Content $Script:ConfigPath -Raw | ConvertFrom-Json
            
            foreach ($key in $Script:DefaultConfig.Keys) {
                if (-not $Script:Config.PSObject.Properties.Name.Contains($key)) {
                    $Script:Config | Add-Member -MemberType NoteProperty -Name $key -Value $Script:DefaultConfig[$key]
                }
            }
        }
        catch {
            $Script:Config = $Script:DefaultConfig
        }
    }
    else {
        $configDir = Split-Path $Script:ConfigPath -Parent
        if (-not (Test-Path $configDir)) {
            New-Item -Path $configDir -ItemType Directory -Force | Out-Null
        }
        
        $Script:Config = $Script:DefaultConfig
        try {
            Save-Configuration
        }
        catch {}
    }
    
    if (-not (Test-Path $Script:Config.LoggingPath)) {
        try {
            New-Item -Path $Script:Config.LoggingPath -ItemType Directory -Force | Out-Null
        }
        catch {}
    }
}

function Save-Configuration {
    try {
        $Script:Config | ConvertTo-Json -Depth 10 | Set-Content $Script:ConfigPath -Force
        Write-Log "Configuration saved" -Level Info
    }
    catch {
        Write-Log "Failed to save configuration: $_" -Level Error
    }
}

# ============================================================================
# LOGGING AND AUDIT
# ============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # File output only - no console spam
    if ($Script:Config -and $Script:Config.LoggingPath) {
        try {
            $logFile = Join-Path $Script:Config.LoggingPath "AccessManager_$(Get-Date -Format 'yyyy-MM-dd').log"
            Add-Content -Path $logFile -Value $logMessage -ErrorAction SilentlyContinue
        }
        catch {}
    }
}

function Write-AuditLog {
    param(
        [Parameter(Mandatory)]
        [string]$Action,
        
        [Parameter()]
        [int]$TargetTechId,
        
        [Parameter()]
        [string]$TargetUserName,
        
        [Parameter()]
        [hashtable]$Details,
        
        [Parameter()]
        [string]$Result = "success",
        
        [Parameter()]
        [string]$Message = ""
    )
    
    if (-not $Script:Config.EnableAuditLog) {
        return
    }
    
    $auditEntry = @{
        timestamp = (Get-Date).ToString("o")
        actor = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        platform = $Script:CurrentPlatform
        db_path = $Script:CurrentDatabasePath
        connection_mode = $Script:CurrentConnectionMode
        action = $Action
        target_techid = $TargetTechId
        target_username = $TargetUserName
        details = $Details
        result = $Result
        message = $Message
        whatif = $Script:WhatIfMode
        readonly = $Script:ReadOnlyMode
    }
    
    try {
        $auditFile = Join-Path $Script:Config.LoggingPath "audit_$(Get-Date -Format 'yyyy-MM-dd').jsonl"
        $auditEntry | ConvertTo-Json -Compress -Depth 10 | Add-Content -Path $auditFile -ErrorAction Stop
    }
    catch {
        Write-Log "Failed to write audit log: $_" -Level Warning
    }
}

# ============================================================================
# DATABASE CONNECTION
# ============================================================================

function New-ODBCConnection {
    param(
        [Parameter(Mandatory)]
        [string]$DSN,
        
        [switch]$ReadOnly
    )
    
    $connString = "DSN=$DSN;Trusted_Connection=Yes;"
    
    try {
        $connection = New-Object System.Data.Odbc.OdbcConnection($connString)
        $connection.Open()
        
        Write-Log "Connected to DSN: $DSN, Database: $($connection.Database)" -Level Success
        
        return $connection
    }
    catch {
        $errorMsg = $_.Exception.Message
        
        if ($errorMsg -match "Data source name not found") {
            throw "DSN '$DSN' not found. Please verify:`n`n1. Open 'ODBC Data Sources (64-bit)'`n2. Check System DSN tab for '$DSN'`n3. Verify it's configured correctly`n`nOriginal error: $errorMsg"
        }
        elseif ($errorMsg -match "Login failed" -or $errorMsg -match "Cannot open database") {
            throw "Authentication or database access failed for DSN '$DSN'.`n`nCheck:`n- Windows Authentication is enabled`n- You have permissions to the database`n- Database exists on server`n`nOriginal error: $errorMsg"
        }
        else {
            throw "Failed to connect via ODBC DSN '$DSN': $errorMsg"
        }
    }
}

function Invoke-DatabaseQuery {
    param(
        [Parameter(Mandatory)]
        $Connection,
        
        [Parameter(Mandatory)]
        [string]$Query,
        
        [Parameter()]
        [hashtable]$Parameters = @{},
        
        [Parameter()]
        [switch]$NonQuery,
        
        [Parameter()]
        $Transaction
    )
    
    if ($Script:WhatIfMode -and ($Query -match "^(INSERT|UPDATE|DELETE)")) {
        Write-Log "WHATIF: $Query" -Level Info
        Write-Log "WHATIF Parameters: $($Parameters | ConvertTo-Json -Compress)" -Level Info
        return $null
    }
    
    try {
        $command = $Connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 30
        
        if ($Transaction) {
            $command.Transaction = $Transaction
        }
        
        foreach ($paramName in $Parameters.Keys) {
            $param = $command.CreateParameter()
            $param.ParameterName = $paramName
            $param.Value = if ($null -eq $Parameters[$paramName]) { [DBNull]::Value } else { $Parameters[$paramName] }
            [void]$command.Parameters.Add($param)
        }
        
        if ($NonQuery) {
            return $command.ExecuteNonQuery()
        }
        else {
            $adapter = New-Object System.Data.Odbc.OdbcDataAdapter($command)
            $dataTable = New-Object System.Data.DataTable
            [void]$adapter.Fill($dataTable)
            return $dataTable
        }
    }
    catch {
        Write-Log "Query failed: $Query | Error: $_" -Level Error
        throw
    }
    finally {
        if ($command) {
            $command.Dispose()
        }
    }
}

function Get-TableNames {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("PRIMARY", "SECONDARY")]
        [string]$Platform
    )
    
    # Based on the Linked Table Manager screenshot
    # Tables are prefixed with "dbo.tb_"
    return @{
        TechTable = "dbo.tb_Technician"
        PrivTable = "dbo.tb_TechPrivilege"
        VerifyTable = "dbo.tb_TechVerify"
        LabOpsLiveTable = "dbo.tb_Technician"  # LabOps live table
    }
}

function Test-TableExists {
    param(
        [Parameter(Mandatory)]
        $Connection,
        
        [Parameter(Mandatory)]
        [string]$TableName
    )
    
    try {
        $testQuery = "SELECT TOP 1 * FROM $TableName"
        $result = Invoke-DatabaseQuery -Connection $Connection -Query $testQuery
        return $true
    }
    catch {
        Write-Log "Table $TableName does not exist or is not accessible" -Level Warning
        return $false
    }
}

# ============================================================================
# HELPER FUNCTION: Convert DataTable to PSObject Array
# ============================================================================

function ConvertFrom-DataTable {
    param(
        [Parameter(Mandatory)]
        [System.Data.DataTable]$DataTable
    )
    
    $results = New-Object System.Collections.ArrayList
    
    foreach ($row in $DataTable.Rows) {
        $obj = New-Object PSObject
        
        foreach ($column in $DataTable.Columns) {
            $columnName = $column.ColumnName
            $value = $row[$columnName]
            
            # Handle DBNull
            if ($value -is [System.DBNull]) {
                $value = $null
            }
            
            $obj | Add-Member -MemberType NoteProperty -Name $columnName -Value $value
        }
        
        [void]$results.Add($obj)
    }
    
    return $results
}

# ============================================================================
# BUSINESS LOGIC
# ============================================================================

function Get-PrivilegesForRole {
    param(
        [Parameter(Mandatory)]
        [string]$Role
    )
    
    $privilegeMap = @{
        "Lab Associate" = @("VIEW")
        "Certified Lab Tech" = @("VIEW", "APPROVE", "ACCEPT")
        "CLS" = @("VIEW", "APPROVE", "ACCEPT")
        "Director" = @("VIEW", "APPROVE", "ACCEPT", "INTERPRETATION", "REAPPROVE")
    }
    
    if ($privilegeMap.ContainsKey($Role)) {
        return $privilegeMap[$Role]
    }
    else {
        Write-Log "Unknown role: $Role. Defaulting to VIEW only." -Level Warning
        return @("VIEW")
    }
}

function Assert-VaxcomEligibility {
    param(
        [Parameter(Mandatory)]
        [string]$JobTitle,
        
        [Parameter(Mandatory)]
        [bool]$IsCLSYes,
        
        [Parameter()]
        [string[]]$Privileges = @()
    )
    
    if ($JobTitle -eq "Lab Associate" -and -not $IsCLSYes) {
        $opsportalPrivileges = @("ACCEPT", "APPROVE")
        $hasVaxcomPriv = $Privileges | Where-Object { $_ -in $opsportalPrivileges }
        
        if ($hasVaxcomPriv) {
            throw "POLICY VIOLATION: Lab Associates without CLS certification cannot have OpsPortal access (ACCEPT/APPROVE privileges)."
        }
        
        Write-Log "OpsPortal eligibility check: Lab Associate without CLS - OK for VIEW only" -Level Info
        return $false
    }
    
    Write-Log "OpsPortal eligibility check: Passed for $JobTitle (CLS: $IsCLSYes)" -Level Info
    return $true
}

function Test-PrivilegeValid {
    param(
        [Parameter(Mandatory)]
        [string]$Privilege
    )
    
    return $Privilege -in $Script:Config.AllowedPrivileges
}

# ============================================================================
# DATA ACCESS LAYER - FIXED
# ============================================================================

function Search-Techs {
    param(
        [Parameter(Mandatory)]
        $Connection,
        
        [Parameter(Mandatory)]
        [string]$Platform,
        
        [Parameter()]
        [string]$UserName,
        
        [Parameter()]
        [int]$TechID,
        
        [Parameter()]
        [string]$Area,
        
        [Parameter()]
        [string]$FirstName,
        
        [Parameter()]
        [string]$LastName,
        
        [Parameter()]
        [string]$EmplID,
        
        [Parameter()]
        [string]$HasPrivilege,
        
        [Parameter()]
        [int]$MaxResults = 1000
    )
    
    $tables = Get-TableNames -Platform $Platform
    $whereClauses = @()
    
    # Build WHERE clause
    if ($UserName) {
        $whereClauses += "T.UserName LIKE '%$UserName%'"
    }
    
    if ($TechID -gt 0) {
        $whereClauses += "T.TechID = $TechID"
    }
    
    if ($Area) {
        $whereClauses += "T.Area LIKE '%$Area%'"
    }
    
    if ($FirstName) {
        $whereClauses += "V.FirstName LIKE '%$FirstName%'"
    }
    
    if ($LastName) {
        $whereClauses += "V.LastName LIKE '%$LastName%'"
    }
    
    if ($EmplID) {
        $whereClauses += "V.emplID LIKE '%$EmplID%'"
    }
    
    if ($HasPrivilege) {
        $whereClauses += "EXISTS (SELECT 1 FROM $($tables.PrivTable) P WHERE P.TechID = T.TechID AND P.Privilege = '$HasPrivilege')"
    }
    
    $whereClause = if ($whereClauses.Count -gt 0) { "WHERE " + ($whereClauses -join " AND ") } else { "" }
    
    $query = @"
SELECT TOP $MaxResults
    T.TechID,
    T.UserName,
    T.Area,
    T.Name,
    T.Password,
    T.Multi,
    V.FirstName,
    V.LastName,
    V.emplID
FROM $($tables.TechTable) T
LEFT JOIN $($tables.VerifyTable) V ON T.TechID = V.TechID
$whereClause
ORDER BY T.UserName
"@
    
    try {
        Write-Log "Executing search query: $query" -Level Info
        
        $results = Invoke-DatabaseQuery -Connection $Connection -Query $query
        
        if (-not $results) {
            Write-Log "Query returned null result" -Level Warning
            return New-Object System.Collections.ArrayList
        }
        
        Write-Log "Query returned $($results.Rows.Count) rows" -Level Info
        
        # Convert DataTable to PSObject array
        $outputList = New-Object System.Collections.ArrayList
        
        foreach ($row in $results.Rows) {
            # Safely extract values with null handling
            $techId = if ($row["TechID"] -is [DBNull]) { 0 } else { [int]$row["TechID"] }
            $userName = if ($row["UserName"] -is [DBNull]) { "" } else { $row["UserName"].ToString().Trim() }
            $area = if ($row["Area"] -is [DBNull]) { "" } else { $row["Area"].ToString().Trim() }
            $name = if ($row["Name"] -is [DBNull]) { "" } else { $row["Name"].ToString().Trim() }
            $password = if ($row["Password"] -is [DBNull]) { "" } else { $row["Password"].ToString().Trim() }
            $multi = if ($row["Multi"] -is [DBNull]) { "" } else { $row["Multi"].ToString().Trim() }
            $firstName = if ($row["FirstName"] -is [DBNull]) { "" } else { $row["FirstName"].ToString().Trim() }
            $lastName = if ($row["LastName"] -is [DBNull]) { "" } else { $row["LastName"].ToString().Trim() }
            $emplID = if ($row["emplID"] -is [DBNull]) { "" } else { $row["emplID"].ToString().Trim() }
            
            # Get privileges for this tech
            $privQuery = "SELECT Privilege FROM $($tables.PrivTable) WHERE TechID = $techId"
            
            $privileges = ""
            try {
                $privResults = Invoke-DatabaseQuery -Connection $Connection -Query $privQuery
                
                if ($privResults -and $privResults.Rows.Count -gt 0) {
                    $privList = @()
                    foreach ($privRow in $privResults.Rows) {
                        if ($privRow["Privilege"] -isnot [DBNull]) {
                            $privList += $privRow["Privilege"].ToString().Trim()
                        }
                    }
                    $privileges = $privList -join ", "
                }
            }
            catch {
                Write-Log "Failed to get privileges for TechID $techId : $_" -Level Warning
                $privileges = ""
            }
            
            # Create PSCustomObject with proper type handling
            $obj = [PSCustomObject]@{
                TechID = $techId
                UserName = $userName
                Area = $area
                Name = $name
                Password = $password
                Multi = $multi
                FirstName = $firstName
                LastName = $lastName
                emplID = $emplID
                Privileges = $privileges
            }
            
            [void]$outputList.Add($obj)
        }
        
        Write-Log "Processed $($outputList.Count) user records" -Level Success
        return $outputList
    }
    catch {
        Write-Log "Search failed: $_" -Level Error
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
        throw
    }
}

function Read-UserFull {
    param(
        [Parameter(Mandatory)]
        $Connection,
        
        [Parameter(Mandatory)]
        [string]$Platform,
        
        [Parameter(Mandatory)]
        [int]$TechID
    )
    
    $tables = Get-TableNames -Platform $Platform
    $user = @{}
    
    try {
        $techQuery = "SELECT * FROM $($tables.TechTable) WHERE TechID = $TechID"
        $techData = Invoke-DatabaseQuery -Connection $Connection -Query $techQuery
        
        if ($techData.Rows.Count -eq 0) {
            throw "TechID $TechID not found"
        }
        
        # Convert to PSObject for easier access
        $techObj = New-Object PSObject
        foreach ($column in $techData.Columns) {
            $value = $techData.Rows[0][$column.ColumnName]
            if ($value -is [DBNull]) { $value = $null }
            $techObj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $value
        }
        $user.Technician = $techObj
        
        $privQuery = "SELECT Privilege FROM $($tables.PrivTable) WHERE TechID = $TechID"
        $privData = Invoke-DatabaseQuery -Connection $Connection -Query $privQuery
        $privList = @()
        foreach ($row in $privData.Rows) {
            if ($row["Privilege"] -isnot [DBNull]) {
                $privList += $row["Privilege"].ToString()
            }
        }
        $user.Privileges = $privList
        
        $verifyQuery = "SELECT * FROM $($tables.VerifyTable) WHERE TechID = $TechID"
        $verifyData = Invoke-DatabaseQuery -Connection $Connection -Query $verifyQuery
        
        if ($verifyData.Rows.Count -gt 0) {
            $verifyObj = New-Object PSObject
            foreach ($column in $verifyData.Columns) {
                $value = $verifyData.Rows[0][$column.ColumnName]
                if ($value -is [DBNull]) { $value = $null }
                $verifyObj | Add-Member -MemberType NoteProperty -Name $column.ColumnName -Value $value
            }
            $user.Verify = $verifyObj
        }
        else {
            $user.Verify = $null
        }
        
        return $user
    }
    catch {
        Write-Log "Failed to read user $TechID : $_" -Level Error
        throw
    }
}

function Grant-Privilege {
    param(
        [Parameter(Mandatory)]
        $Connection,
        
        [Parameter(Mandatory)]
        [string]$Platform,
        
        [Parameter(Mandatory)]
        [int]$TechID,
        
        [Parameter(Mandatory)]
        [string]$Privilege
    )
    
    if (-not (Test-PrivilegeValid -Privilege $Privilege)) {
        throw "INVALID PRIVILEGE: '$Privilege' is not in the allowed list"
    }
    
    $tables = Get-TableNames -Platform $Platform
    
    $dupCheck = "SELECT COUNT(*) as Cnt FROM $($tables.PrivTable) WHERE TechID = $TechID AND Privilege = '$Privilege'"
    $dupResult = Invoke-DatabaseQuery -Connection $Connection -Query $dupCheck
    
    if ($dupResult.Rows[0]["Cnt"] -gt 0) {
        Write-Log "Privilege '$Privilege' already granted to TechID $TechID" -Level Warning
        return
    }
    
    $userInfo = Read-UserFull -Connection $Connection -Platform $Platform -TechID $TechID
    
    try {
        $insertPriv = "INSERT INTO $($tables.PrivTable) (TechID, Privilege) VALUES ($TechID, '$Privilege')"
        
        Invoke-DatabaseQuery -Connection $Connection -Query $insertPriv -NonQuery
        
        Write-AuditLog -Action "grant" -TargetTechId $TechID -TargetUserName $userInfo.Technician.UserName -Details @{
            privilege = $Privilege
        } -Result "success" -Message "Privilege granted"
        
        Write-Log "Privilege '$Privilege' granted to TechID $TechID" -Level Success
    }
    catch {
        Write-Log "Failed to grant privilege: $_" -Level Error
        throw
    }
}

function Revoke-Privilege {
    param(
        [Parameter(Mandatory)]
        $Connection,
        
        [Parameter(Mandatory)]
        [string]$Platform,
        
        [Parameter(Mandatory)]
        [int]$TechID,
        
        [Parameter(Mandatory)]
        [string]$Privilege
    )
    
    $tables = Get-TableNames -Platform $Platform
    $userInfo = Read-UserFull -Connection $Connection -Platform $Platform -TechID $TechID
    
    try {
        $deletePriv = "DELETE FROM $($tables.PrivTable) WHERE TechID = $TechID AND Privilege = '$Privilege'"
        
        $rowsAffected = Invoke-DatabaseQuery -Connection $Connection -Query $deletePriv -NonQuery
        
        Write-AuditLog -Action "revoke" -TargetTechId $TechID -TargetUserName $userInfo.Technician.UserName -Details @{
            privilege = $Privilege
        } -Result "success" -Message "Privilege revoked"
        
        Write-Log "Privilege '$Privilege' revoked from TechID $TechID" -Level Success
    }
    catch {
        Write-Log "Failed to revoke privilege: $_" -Level Error
        throw
    }
}

# ============================================================================
# WPF USER INTERFACE
# ============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Access Manager - LABOPS/OpsPortal RBAC System" 
        Height="800" Width="1400"
        WindowStartupLocation="CenterScreen"
        Background="#F5F5F5">
    
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Background" Value="#2196F3"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Height" Value="30"/>
        </Style>
        
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Height" Value="30"/>
        </Style>
        
        <Style TargetType="Label">
            <Setter Property="Margin" Value="5,5,5,0"/>
            <Setter Property="FontWeight" Value="Medium"/>
        </Style>
        
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#F44336"/>
        </Style>
        
        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#4CAF50"/>
        </Style>
    </Window.Resources>
    
    <DockPanel>
        <!-- Top Menu Bar -->
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem Name="MenuFileExit" Header="E_xit"/>
            </MenuItem>
            <MenuItem Header="_Tools">
                <MenuItem Name="MenuToolsWhatIf" Header="_WhatIf Mode" IsCheckable="True"/>
                <MenuItem Name="MenuToolsReadOnly" Header="_Read-Only Mode" IsCheckable="True"/>
            </MenuItem>
            <MenuItem Header="_Help">
                <MenuItem Name="MenuHelpAbout" Header="_About"/>
            </MenuItem>
        </Menu>
        
        <!-- Status Bar -->
        <StatusBar DockPanel.Dock="Bottom" Height="30" Background="#E0E0E0">
            <StatusBarItem>
                <TextBlock Name="StatusText" Text="Ready" FontWeight="Medium"/>
            </StatusBarItem>
            <Separator/>
            <StatusBarItem>
                <TextBlock Name="StatusDatabase" Text="Not Connected"/>
            </StatusBarItem>
            <Separator/>
            <StatusBarItem>
                <TextBlock Name="StatusPlatform" Text="Platform: None"/>
            </StatusBarItem>
            <Separator/>
            <StatusBarItem>
                <TextBlock Name="StatusRowCount" Text="Rows: 0"/>
            </StatusBarItem>
            <Separator/>
            <StatusBarItem HorizontalAlignment="Right">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Name="StatusWhatIf" Text="" Foreground="Orange" FontWeight="Bold" Margin="0,0,10,0"/>
                    <TextBlock Name="StatusReadOnly" Text="" Foreground="Red" FontWeight="Bold"/>
                </StackPanel>
            </StatusBarItem>
        </StatusBar>
        
        <!-- Main Content Area -->
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Connection Panel -->
            <Border Grid.Row="0" Background="White" BorderBrush="#CCCCCC" BorderThickness="0,0,0,1" Padding="10">
                <StackPanel>
                    <Label Content="Database Connection (ODBC/DSN Mode)" FontSize="16" FontWeight="Bold"/>
                    
                    <Grid Margin="0,10,0,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="200"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        
                        <StackPanel Grid.Column="0" Orientation="Vertical">
                            <Label Content="Platform:" Margin="0"/>
                            <ComboBox Name="CmbPlatform" Height="30" Margin="0">
                                <ComboBoxItem Content="PRIMARY (DSN: LAB_PRIMARY)" Tag="PRIMARY"/>
                                <ComboBoxItem Content="SECONDARY (DSN: LAB_SECONDARY)" Tag="SECONDARY"/>
                            </ComboBox>
                        </StackPanel>
                        
                        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom">
                            <Button Name="BtnConnect" Content="Connect" Width="100" Style="{StaticResource SuccessButton}"/>
                            <Button Name="BtnDisconnect" Content="Disconnect" Width="100" IsEnabled="False"/>
                        </StackPanel>
                    </Grid>
                </StackPanel>
            </Border>
            
            <!-- Tabbed Interface -->
            <TabControl Grid.Row="1" Margin="10" Name="MainTabs">
                
                <!-- MANAGE TAB -->
                <TabItem Header="Manage Users">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Search Panel -->
                        <Border Grid.Row="0" Background="White" Padding="10" Margin="0,0,0,10">
                            <StackPanel>
                                <Label Content="Search Technicians" FontSize="14" FontWeight="Bold"/>
                                
                                <Grid Margin="0,10,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    
                                    <StackPanel Grid.Row="0" Grid.Column="0">
                                        <Label Content="UserName:"/>
                                        <TextBox Name="TxtSearchUserName"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="0" Grid.Column="1">
                                        <Label Content="TechID:"/>
                                        <TextBox Name="TxtSearchTechID"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="0" Grid.Column="2">
                                        <Label Content="Area:"/>
                                        <TextBox Name="TxtSearchArea"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="0" Grid.Column="3">
                                        <Label Content="First Name:"/>
                                        <TextBox Name="TxtSearchFirstName"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="1" Grid.Column="0">
                                        <Label Content="Last Name:"/>
                                        <TextBox Name="TxtSearchLastName"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="1" Grid.Column="1">
                                        <Label Content="Employee ID:"/>
                                        <TextBox Name="TxtSearchEmplID"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="1" Grid.Column="2">
                                        <Label Content="Has Privilege:"/>
                                        <ComboBox Name="CmbSearchPrivilege"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="1" Grid.Column="3" VerticalAlignment="Bottom">
                                        <Button Name="BtnSearch" Content="🔍 Search" Height="35" Margin="5,20,5,5"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="4" Orientation="Horizontal" HorizontalAlignment="Right">
                                        <Button Name="BtnClearSearch" Content="Clear Filters" Width="100"/>
                                        <Button Name="BtnRefresh" Content="↻ Refresh" Width="100"/>
                                        <Button Name="BtnExportGrid" Content="📄 Export CSV" Width="120"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>
                        </Border>
                        
                        <!-- Data Grid -->
                        <DataGrid Grid.Row="1" Name="DataGridUsers" 
                                  AutoGenerateColumns="False" 
                                  IsReadOnly="True"
                                  SelectionMode="Single"
                                  CanUserAddRows="False"
                                  CanUserDeleteRows="False"
                                  AlternatingRowBackground="#F9F9F9"
                                  GridLinesVisibility="Horizontal"
                                  HeadersVisibility="Column">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="TechID" Binding="{Binding TechID}" Width="70"/>
                                <DataGridTextColumn Header="UserName" Binding="{Binding UserName}" Width="120"/>
                                <DataGridTextColumn Header="Area" Binding="{Binding Area}" Width="100"/>
                                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="180"/>
                                <DataGridTextColumn Header="First Name" Binding="{Binding FirstName}" Width="120"/>
                                <DataGridTextColumn Header="Last Name" Binding="{Binding LastName}" Width="120"/>
                                <DataGridTextColumn Header="EmplID" Binding="{Binding emplID}" Width="100"/>
                                <DataGridTextColumn Header="Privileges" Binding="{Binding Privileges}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        
                        <!-- Action Buttons -->
                        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                            <Button Name="BtnEditPrivileges" Content="✏️ Edit Privileges" Width="150" IsEnabled="False"/>
                            <Button Name="BtnViewDetails" Content="👁️ View Details" Width="150" IsEnabled="False"/>
                        </StackPanel>
                    </Grid>
                </TabItem>
                
                <!-- PRIVILEGES TAB -->
                <TabItem Header="Manage Privileges" Name="TabPrivileges">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <Border Grid.Row="0" Background="White" Padding="20" Margin="0,0,0,10">
                            <StackPanel>
                                <Label Content="Selected User" FontSize="14" FontWeight="Bold"/>
                                <TextBlock Name="TxtSelectedUser" FontSize="12" Foreground="Gray" Margin="5"/>
                                
                                <Label Content="Current Privileges:" Margin="0,10,0,0"/>
                                <TextBlock Name="TxtCurrentPrivileges" FontSize="12" Margin="5"/>
                            </StackPanel>
                        </Border>
                        
                        <Border Grid.Row="1" Background="White" Padding="20">
                            <StackPanel>
                                <Label Content="Grant/Revoke Privilege" FontSize="14" FontWeight="Bold"/>
                                
                                <Grid Margin="0,10,0,0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    
                                    <StackPanel Grid.Column="0">
                                        <Label Content="Select Privilege:"/>
                                        <ComboBox Name="CmbPrivilege" Height="35"/>
                                    </StackPanel>
                                    
                                    <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom">
                                        <Button Name="BtnGrantPrivilege" Content="✅ Grant" Width="100" Style="{StaticResource SuccessButton}"/>
                                        <Button Name="BtnRevokePrivilege" Content="❌ Revoke" Width="100" Style="{StaticResource DangerButton}"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>
                        </Border>
                    </Grid>
                </TabItem>
                
                <!-- LOGS TAB -->
                <TabItem Header="Console / Logs">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="10">
                            <Button Name="BtnClearConsole" Content="Clear Console" Width="120"/>
                            <Button Name="BtnOpenLogFolder" Content="Open Log Folder" Width="140" Margin="5,0"/>
                        </StackPanel>
                        
                        <TextBox Grid.Row="1" Name="TxtConsole" 
                                 IsReadOnly="True"
                                 Background="Black"
                                 Foreground="LightGreen"
                                 FontFamily="Consolas"
                                 FontSize="12"
                                 VerticalScrollBarVisibility="Auto"
                                 TextWrapping="Wrap"
                                 Margin="10"/>
                    </Grid>
                </TabItem>
                
            </TabControl>
        </Grid>
    </DockPanel>
</Window>
"@

# ============================================================================
# WPF INITIALIZATION AND EVENT HANDLERS
# ============================================================================

function Initialize-WpfWindow {
    try {
        $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        $controls = @{}
        $xaml | Select-String -Pattern 'Name="(\w+)"' -AllMatches | ForEach-Object {
            $_.Matches | ForEach-Object {
                $name = $_.Groups[1].Value
                $controls[$name] = $window.FindName($name)
            }
        }
        
        return @{
            Window = $window
            Controls = $controls
        }
    }
    catch {
        Write-Log "Failed to create WPF window: $_" -Level Error
        throw
    }
}

function Update-StatusBar {
    param(
        $Controls,
        [string]$Message
    )
    
    if ($Message) {
        $Controls.StatusText.Text = $Message
    }
    
    if ($Script:CurrentDatabasePath) {
        $Controls.StatusDatabase.Text = "DSN: $Script:CurrentDatabasePath"
    }
    else {
        $Controls.StatusDatabase.Text = "Not Connected"
    }
    
    if ($Script:CurrentPlatform) {
        $Controls.StatusPlatform.Text = "Platform: $Script:CurrentPlatform"
    }
    
    if ($Script:WhatIfMode) {
        $Controls.StatusWhatIf.Text = "⚠️ WHATIF MODE"
    }
    else {
        $Controls.StatusWhatIf.Text = ""
    }
    
    if ($Script:ReadOnlyMode) {
        $Controls.StatusReadOnly.Text = "🔒 READ-ONLY"
    }
    else {
        $Controls.StatusReadOnly.Text = ""
    }
}

function Register-EventHandlers {
    param(
        $Window,
        $Controls
    )
    
    # Menu Events
    $Controls.MenuFileExit.Add_Click({
        $Window.Close()
    })
    
    $Controls.MenuToolsWhatIf.Add_Checked({
        $Script:WhatIfMode = $true
        Update-StatusBar -Controls $Controls -Message "WhatIf Mode ENABLED"
        Write-Log "WhatIf Mode enabled" -Level Warning
    })
    
    $Controls.MenuToolsWhatIf.Add_Unchecked({
        $Script:WhatIfMode = $false
        Update-StatusBar -Controls $Controls -Message "WhatIf Mode disabled"
        Write-Log "WhatIf Mode disabled" -Level Info
    })
    
    $Controls.MenuToolsReadOnly.Add_Checked({
        $Script:ReadOnlyMode = $true
        Update-StatusBar -Controls $Controls -Message "Read-Only Mode ENABLED"
        Write-Log "Read-Only Mode enabled" -Level Warning
    })
    
    $Controls.MenuToolsReadOnly.Add_Unchecked({
        $Script:ReadOnlyMode = $false
        Update-StatusBar -Controls $Controls -Message "Read-Only Mode disabled"
        Write-Log "Read-Only Mode disabled" -Level Info
    })
    
    $Controls.MenuHelpAbout.Add_Click({
        [System.Windows.MessageBox]::Show("Access Manager v$Script:AppVersion`n`nLABOPS/OpsPortal RBAC Management System`n`nConnects via ODBC to:`n- PRIMARY: DSN LAB_PRIMARY`n- SECONDARY: DSN LAB_SECONDARY`n`nLast Updated: 2025-10-29", "About", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    })
    
    # Connect Button
    $Controls.BtnConnect.Add_Click({
        $platform = $Controls.CmbPlatform.SelectedItem.Tag
        
        if ([string]::IsNullOrWhiteSpace($platform)) {
            [System.Windows.MessageBox]::Show("Please select a platform (PRIMARY or SECONDARY).", "Validation Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $dsn = if ($platform -eq "PRIMARY") { $Script:Config.ODBCDSN_PRIMARY } else { $Script:Config.ODBCDSN_SECONDARY }
        
        try {
            Update-StatusBar -Controls $Controls -Message "Connecting to $dsn..."
            
            $conn = New-ODBCConnection -DSN $dsn
            
            # Verify tables exist
            $tables = Get-TableNames -Platform $platform
            $missingTables = @()
            
            foreach ($tableName in @($tables.TechTable, $tables.PrivTable, $tables.VerifyTable)) {
                if (-not (Test-TableExists -Connection $conn -TableName $tableName)) {
                    $missingTables += $tableName
                }
            }
            
            if ($missingTables.Count -gt 0) {
                throw "Required tables not found: $($missingTables -join ', ')`n`nPlease verify:`n1. Correct DSN selected`n2. Database accessible`n3. Tables exist on server"
            }
            
            $Script:CurrentConnection = $conn
            $Script:CurrentDatabasePath = $dsn
            $Script:CurrentPlatform = $platform
            $Script:CurrentConnectionMode = "ODBC"
            
            $Controls.BtnConnect.IsEnabled = $false
            $Controls.BtnDisconnect.IsEnabled = $true
            $Controls.BtnSearch.IsEnabled = $true
            
            # Populate privilege dropdown
            $Controls.CmbSearchPrivilege.Items.Clear()
            $Controls.CmbSearchPrivilege.Items.Add("") | Out-Null
            foreach ($priv in $Script:Config.AllowedPrivileges) {
                $Controls.CmbSearchPrivilege.Items.Add($priv) | Out-Null
            }
            
            $Controls.CmbPrivilege.Items.Clear()
            foreach ($priv in $Script:Config.AllowedPrivileges) {
                $Controls.CmbPrivilege.Items.Add($priv) | Out-Null
            }
            
            Update-StatusBar -Controls $Controls -Message "Connected successfully to $dsn"
            
            # Auto-load all users
            Invoke-SearchUsers -Controls $Controls
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to connect to database:`n`n$_", "Connection Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Connection failed: $_" -Level Error
        }
    })
    
    # Disconnect Button
    $Controls.BtnDisconnect.Add_Click({
        if ($Script:CurrentConnection) {
            try {
                $Script:CurrentConnection.Close()
                $Script:CurrentConnection.Dispose()
            }
            catch {}
            
            $Script:CurrentConnection = $null
            $Script:CurrentDatabasePath = $null
            $Script:CurrentPlatform = $null
            
            $Controls.BtnConnect.IsEnabled = $true
            $Controls.BtnDisconnect.IsEnabled = $false
            $Controls.DataGridUsers.ItemsSource = $null
            
            Update-StatusBar -Controls $Controls -Message "Disconnected"
        }
    })
    
    # Search Button
    $Controls.BtnSearch.Add_Click({
        Invoke-SearchUsers -Controls $Controls
    })
    
    # Clear Search
    $Controls.BtnClearSearch.Add_Click({
        $Controls.TxtSearchUserName.Text = ""
        $Controls.TxtSearchTechID.Text = ""
        $Controls.TxtSearchArea.Text = ""
        $Controls.TxtSearchFirstName.Text = ""
        $Controls.TxtSearchLastName.Text = ""
        $Controls.TxtSearchEmplID.Text = ""
        $Controls.CmbSearchPrivilege.SelectedIndex = 0
        
        Invoke-SearchUsers -Controls $Controls
    })
    
    # Refresh Button
    $Controls.BtnRefresh.Add_Click({
        Invoke-SearchUsers -Controls $Controls
    })
    
    # DataGrid Selection
    $Controls.DataGridUsers.Add_SelectionChanged({
        $hasSelection = $Controls.DataGridUsers.SelectedItem -ne $null
        $Controls.BtnEditPrivileges.IsEnabled = $hasSelection
        $Controls.BtnViewDetails.IsEnabled = $hasSelection
        
        if ($hasSelection) {
            $selected = $Controls.DataGridUsers.SelectedItem
            $Controls.TxtSelectedUser.Text = "TechID: $($selected.TechID) | UserName: $($selected.UserName) | Name: $($selected.Name)"
            $Controls.TxtCurrentPrivileges.Text = if ($selected.Privileges) { $selected.Privileges } else { "(none)" }
        }
    })
    
    # Edit Privileges Button
    $Controls.BtnEditPrivileges.Add_Click({
        $selected = $Controls.DataGridUsers.SelectedItem
        if ($selected) {
            $Controls.MainTabs.SelectedIndex = 1  # Switch to Privileges tab
        }
    })
    
    # View Details Button
    $Controls.BtnViewDetails.Add_Click({
        $selected = $Controls.DataGridUsers.SelectedItem
        if ($selected) {
            try {
                $fullUser = Read-UserFull -Connection $Script:CurrentConnection -Platform $Script:CurrentPlatform -TechID $selected.TechID
                
                $details = @"
User Details
============

Technician Information:
  TechID: $($fullUser.Technician.TechID)
  UserName: $($fullUser.Technician.UserName)
  Name: $($fullUser.Technician.Name)
  Area: $($fullUser.Technician.Area)
  Password: $($fullUser.Technician.Password)

Verification Information:
  First Name: $($fullUser.Verify.FirstName)
  Last Name: $($fullUser.Verify.LastName)
  Employee ID: $($fullUser.Verify.emplID)

Privileges:
  $($fullUser.Privileges -join ', ')
"@
                
                [System.Windows.MessageBox]::Show($details, "User Details", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to load user details: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })
    
    # Grant Privilege Button
    $Controls.BtnGrantPrivilege.Add_Click({
        $selected = $Controls.DataGridUsers.SelectedItem
        $privilege = $Controls.CmbPrivilege.SelectedItem
        
        if (-not $selected) {
            [System.Windows.MessageBox]::Show("Please select a user from the Manage Users tab.", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($privilege)) {
            [System.Windows.MessageBox]::Show("Please select a privilege to grant.", "No Privilege", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $confirm = [System.Windows.MessageBox]::Show("Grant privilege '$privilege' to user '$($selected.UserName)'?", "Confirm Grant", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Grant-Privilege -Connection $Script:CurrentConnection -Platform $Script:CurrentPlatform -TechID $selected.TechID -Privilege $privilege
                [System.Windows.MessageBox]::Show("Privilege granted successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                Invoke-SearchUsers -Controls $Controls
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to grant privilege: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })
    
    # Revoke Privilege Button
    $Controls.BtnRevokePrivilege.Add_Click({
        $selected = $Controls.DataGridUsers.SelectedItem
        $privilege = $Controls.CmbPrivilege.SelectedItem
        
        if (-not $selected) {
            [System.Windows.MessageBox]::Show("Please select a user from the Manage Users tab.", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($privilege)) {
            [System.Windows.MessageBox]::Show("Please select a privilege to revoke.", "No Privilege", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $confirm = [System.Windows.MessageBox]::Show("⚠️ REVOKE privilege '$privilege' from user '$($selected.UserName)'?`n`nThis action will be logged.", "Confirm Revoke", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        
        if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
            try {
                Revoke-Privilege -Connection $Script:CurrentConnection -Platform $Script:CurrentPlatform -TechID $selected.TechID -Privilege $privilege
                [System.Windows.MessageBox]::Show("Privilege revoked successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                Invoke-SearchUsers -Controls $Controls
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to revoke privilege: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })
    
    # Export Grid
    $Controls.BtnExportGrid.Add_Click({
        if ($Controls.DataGridUsers.ItemsSource) {
            try {
                $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
                $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
                $saveDialog.FileName = "AccessManager_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                
                if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $Controls.DataGridUsers.ItemsSource | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
                    [System.Windows.MessageBox]::Show("Data exported successfully to:`n$($saveDialog.FileName)", "Export Complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export data: $_", "Export Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            }
        }
    })
    
    # Console
    $Controls.BtnClearConsole.Add_Click({
        $Controls.TxtConsole.Text = ""
    })
    
    $Controls.BtnOpenLogFolder.Add_Click({
        if (Test-Path $Script:Config.LoggingPath) {
            Start-Process "explorer.exe" -ArgumentList $Script:Config.LoggingPath
        }
    })
}

function Invoke-SearchUsers {
    param(
        $Controls
    )
    
    if (-not $Script:CurrentConnection) {
        return
    }
    
    try {
        $searchParams = @{
            Connection = $Script:CurrentConnection
            Platform = $Script:CurrentPlatform
        }
        
        if ($Controls.TxtSearchUserName.Text) {
            $searchParams.UserName = $Controls.TxtSearchUserName.Text
        }
        
        if ($Controls.TxtSearchTechID.Text -match '^\d+$') {
            $searchParams.TechID = [int]$Controls.TxtSearchTechID.Text
        }
        
        if ($Controls.TxtSearchArea.Text) {
            $searchParams.Area = $Controls.TxtSearchArea.Text
        }
        
        if ($Controls.TxtSearchFirstName.Text) {
            $searchParams.FirstName = $Controls.TxtSearchFirstName.Text
        }
        
        if ($Controls.TxtSearchLastName.Text) {
            $searchParams.LastName = $Controls.TxtSearchLastName.Text
        }
        
        if ($Controls.TxtSearchEmplID.Text) {
            $searchParams.EmplID = $Controls.TxtSearchEmplID.Text
        }
        
        if ($Controls.CmbSearchPrivilege.SelectedItem -and $Controls.CmbSearchPrivilege.SelectedItem -ne "") {
            $searchParams.HasPrivilege = $Controls.CmbSearchPrivilege.SelectedItem
        }
        
        $results = Search-Techs @searchParams
        
        # Results is now an ArrayList of PSObjects, can bind directly to DataGrid
        $Controls.DataGridUsers.ItemsSource = $results
        
        $rowCount = if ($results) { $results.Count } else { 0 }
        $Controls.StatusRowCount.Text = "Rows: $rowCount"
        Update-StatusBar -Controls $Controls -Message "Search returned $rowCount results"
    }
    catch {
        [System.Windows.MessageBox]::Show("Search failed: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        Write-Log "Search failed: $_" -Level Error
    }
}

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

function Start-AccessManager {
    # Silent initialization
    Initialize-Configuration
    
    $wpf = Initialize-WpfWindow
    $window = $wpf.Window
    $controls = $wpf.Controls
    
    $controls.CmbPlatform.SelectedIndex = 0
    
    Register-EventHandlers -Window $window -Controls $controls
    
    Update-StatusBar -Controls $controls -Message "Ready to connect"
    
    $window.Add_Closing({
        if ($Script:CurrentConnection) {
            try {
                $Script:CurrentConnection.Close()
                $Script:CurrentConnection.Dispose()
            }
            catch {}
        }
        
        Write-Log "Access Manager closed" -Level Info
    })
    
    $window.ShowDialog() | Out-Null
}

# Start the application
try {
    Start-AccessManager
}
catch {
    Write-Host "FATAL ERROR: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Read-Host "Press Enter to exit"
}