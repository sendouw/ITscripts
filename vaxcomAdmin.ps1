#Requires -Version 5.1
<#
 Access Manager – LABOPS/OpsPortal (PRIMARY/SECONDARY) – WPF GUI (ODBC/DSN)
 Author: Andy Sendouw (tool built with ChatGPT)
 Last Updated: 2025-10-28

 Highlights:
 - WPF single-file, ps2exe-ready
 - ODBC/DSN only (safe for linked tables)
 - Read-Only default; transactions + parameterized writes
 - SOP guardrails:
     * Lab Associate (CLS NO) => block ACCEPT/APPROVE (OpsPortal)
     * CLS YES allows OpsPortal
 - Features:
     * Search/List (DataGrid), view/edit fields
     * Add or Update (upsert)
     * Disable (soft: clear privs, mark INFORMATION/TERMINATED)
     * Mirror privileges from existing TechID
     * Compare two users (A↔B)
     * Create/ensure LabOps live (role)
 - Audit JSONL + rolling logs under %APPDATA%\AccessManager
#>


Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
Add-Type -AssemblyName System.Windows.Forms
[System.Reflection.Assembly]::LoadWithPartialName('System.Data')   | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')| Out-Null
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# --------------------------- Config / Paths ---------------------------
$AppName   = "Access Manager"
$AppVer    = "1.4.0"
$BaseDir   = Join-Path $env:APPDATA "AccessManager"
$LogDir    = Join-Path $BaseDir "logs"
$ConfigFile= Join-Path $BaseDir "config.json"
$AuditFile = Join-Path $LogDir  ("audit_{0}.jsonl" -f (Get-Date -Format 'yyyy-MM-dd'))

if (-not (Test-Path $BaseDir)) { New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null }
if (-not (Test-Path $LogDir))  { New-Item -ItemType Directory -Path $LogDir -Force  | Out-Null }

$DefaultCfg = [ordered]@{
  DefaultPlatform   = "SECONDARY"           # PRIMARY or SECONDARY
  DSN_PRIMARY          = "LAB_PRIMARY"
  DSN_SECONDARY          = "LAB_SECONDARY"
  Database_PRIMARY     = "lab_primary"
  Database_SECONDARY     = "lab_secondary"
  Theme             = "light"
  ReadOnly          = $true
  EnableAudit       = $true
  MaxRows           = 1000
  AllowedPrivileges = @("VIEW","ACCEPT","APPROVE","REAPPROVE","INTERPRETATION","LIMITA","LIMITB")
  AllowedAreas      = @("SPECHEM","ENDO","STEROIDS","TOXICOLOGY","IMMCHEM","BIOCHEMG","THYROID","TIP","TM","LCCORE","SEROLOGY","INFORMATION","MOLMICRO","CLSID","labops")
  SqlSchema        = "dbo"
  TablePrefix      = "tb_"
  ForceSqlNames    = $false   # If true, skip autodetect and force dbo.tb_* names
  TechTableName   = "tb_Technician"
  PrivTableName   = "tb_TechPrivilege"
  VerifyTableName = "tb_TechVerify"
  LabOpsLiveTableName = "tb_LabOpsLive"
  ShowConsoleLog  = $false
}

function Load-Cfg {
  $c = $null
  if (Test-Path $ConfigFile) {
    try { $c = Get-Content $ConfigFile -Raw | ConvertFrom-Json } catch { $c = $null }
  }
  if (-not $c) { $c = $DefaultCfg }
  foreach ($k in $DefaultCfg.Keys) {
    $prop = $c.PSObject.Properties.Match($k)
    if ($prop.Count -eq 0) {
      $c | Add-Member -NotePropertyName $k -NotePropertyValue $DefaultCfg[$k]
      continue
    }
    if ($null -eq $prop[0].Value) {
      $prop[0].Value = $DefaultCfg[$k]
    }
  }
  return $c
}
function Save-Cfg([object]$c) { $c | ConvertTo-Json -Depth 10 | Set-Content -Path $ConfigFile -Force }

function Get-Cfg([string]$Name, $Default) {
  try {
    if ($null -ne $Cfg -and $Cfg.PSObject -and ($Cfg.PSObject.Properties.Match($Name).Count -gt 0)) {
      $val = $Cfg.$Name
      if ($null -ne $val -and $val -ne "") { return $val }
    }
  } catch {}
  return $Default
}

$Cfg = Load-Cfg
try { $null = Get-Tables -Platform $Cfg.DefaultPlatform; Write-Log ("Tables for {0}: {1}" -f $Cfg.DefaultPlatform, ($Script:TableCache[$Cfg.DefaultPlatform] | Out-String)) } catch {}

# --------------------------- Logging / Audit ---------------------------
function Write-Log {
  param([string]$Msg,[string]$Level="INFO")
  $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Msg
  $line | Add-Content -Path (Join-Path $LogDir ("app_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd')))
  $showHost = $false
  try {
    $prop = $Cfg.PSObject.Properties['ShowConsoleLog']
    if ($prop -and [bool]$prop.Value) { $showHost = $true }
  } catch {}

  if ($showHost) {
    Write-Host $line
  }
}
function Audit {
  param([string]$Action,[hashtable]$Details,[string]$Result="success",[string]$Message="")
  if (-not $Cfg.EnableAudit) { return }
  $row = [ordered]@{
    ts      = (Get-Date).ToString("o")
    actor   = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    action  = $Action
    result  = $Result
    message = $Message
    details = $Details
    version = $AppVer
    readonly= $Cfg.ReadOnly
  }
  ($row | ConvertTo-Json -Compress -Depth 12) | Add-Content -Path $AuditFile
}

# --------------------------- DB Helpers (ODBC) ---------------------------
function Get-DSN([ValidateSet("PRIMARY","SECONDARY")]$Platform) {
  if ($Platform -eq "PRIMARY") { $Cfg.DSN_PRIMARY } else { $Cfg.DSN_SECONDARY }
}
# Cache of detected table names per platform
$Script:TableCache = @{}
$Script:ActivePlatform = $Cfg.DefaultPlatform
$Script:LastActionNote = $null
$Script:ActiveTheme = "light"
$Script:ThemePalettes = @{
  light = [ordered]@{
    WindowBackground   = "#FAFBFF"
    Surface            = "#F3F6FB"
    PanelBorder        = "#D6DEED"
    PrimaryAccent      = "#1C3F80"
    PrimaryText        = "#1C3F80"
    SecondaryText      = "#4C5466"
    BodyText           = "#2C3140"
    TipText            = "#6C7285"
    GroupBackground    = "#FFFFFF"
    GridRow            = "#FFFFFF"
    GridAltRow         = "#EDF2FC"
    GridCompareAltRow  = "#F4F7FD"
    GridHeaderBackground = "#E4EAF7"
    GridHeaderForeground = "#1C3F80"
    ControlBackground    = "#FFFFFF"
    ControlBorder        = "#D6DEED"
  }
  dark = [ordered]@{
    WindowBackground   = "#1E2230"
    Surface            = "#2A3042"
    PanelBorder        = "#3C445B"
    PrimaryAccent      = "#7AA7FF"
    PrimaryText        = "#D7E5FF"
    SecondaryText      = "#C3D4FF"
    BodyText           = "#F0F4FF"
    TipText            = "#94A0C6"
    GroupBackground    = "#252B3D"
    GridRow            = "#2A3042"
    GridAltRow         = "#343B51"
    GridCompareAltRow  = "#3A435D"
    GridHeaderBackground = "#343B51"
    GridHeaderForeground = "#D7E5FF"
    ControlBackground    = "#30364A"
    ControlBorder        = "#3C445B"
  }
}

function Update-ResourceBrush([string]$ResourceKey,[string]$ColorHex) {
  if (-not $Window) { return }
  if (-not $Window.Resources.Contains($ResourceKey)) { return }
  $brush = $Window.Resources[$ResourceKey]
  if ($brush -is [System.Windows.Media.SolidColorBrush]) {
    try {
      $color = [System.Windows.Media.ColorConverter]::ConvertFromString($ColorHex)
      if ($color -is [System.Windows.Media.Color]) {
        $brush.Color = $color
      }
    } catch {}
  }
}

function Apply-ThemeByName([string]$ThemeName,[switch]$Persist) {
  if ([string]::IsNullOrWhiteSpace($ThemeName)) { $ThemeName = "light" }
  $key = $ThemeName.ToLowerInvariant()
  if (-not $Script:ThemePalettes.ContainsKey($key)) { $key = "light" }
  if (-not $Window) {
    $Script:ActiveTheme = $key
    if ($Persist) {
      $Cfg.Theme = $key
      Save-Cfg $Cfg
    } else {
      $Cfg.Theme = $key
    }
    return
  }
  $palette = $Script:ThemePalettes[$key]
  Update-ResourceBrush "WindowBackgroundBrush"  $palette.WindowBackground
  Update-ResourceBrush "SurfaceBrush"           $palette.Surface
  Update-ResourceBrush "PanelBorderBrush"       $palette.PanelBorder
  Update-ResourceBrush "PrimaryAccentBrush"     $palette.PrimaryAccent
  Update-ResourceBrush "PrimaryTextBrush"       $palette.PrimaryText
  Update-ResourceBrush "SecondaryTextBrush"     $palette.SecondaryText
  Update-ResourceBrush "BodyTextBrush"          $palette.BodyText
  Update-ResourceBrush "TipTextBrush"           $palette.TipText
  Update-ResourceBrush "GroupBackgroundBrush"   $palette.GroupBackground
  Update-ResourceBrush "GridRowBrush"           $palette.GridRow
  Update-ResourceBrush "GridAltRowBrush"        $palette.GridAltRow
  Update-ResourceBrush "GridCompareAltRowBrush" $palette.GridCompareAltRow
  Update-ResourceBrush "GridHeaderBackgroundBrush" $palette.GridHeaderBackground
  Update-ResourceBrush "GridHeaderForegroundBrush" $palette.GridHeaderForeground
  Update-ResourceBrush "ControlBackgroundBrush"    $palette.ControlBackground
  Update-ResourceBrush "ControlBorderBrush"        $palette.ControlBorder
  try {
    $Window.Background = $Window.Resources["WindowBackgroundBrush"]
  } catch {}
  $Script:ActiveTheme = $key
  if ($Persist) {
    $Cfg.Theme = $key
    Save-Cfg $Cfg
  } else {
    $Cfg.Theme = $key
  }
  Update-DarkModeToggleVisual
}

function Apply-ThemeToAbout([System.Windows.Window]$Dialog) {
  if (-not $Dialog -or -not $Window) { return }
  try {
    if ($Window.Resources.Contains("WindowBackgroundBrush")) {
      $Dialog.Background = $Window.Resources["WindowBackgroundBrush"]
    }
    $badge = $Dialog.FindName('AboutBadgeShape')
    if ($badge -and $Window.Resources.Contains("PrimaryAccentBrush")) {
      $badge.Fill = $Window.Resources["PrimaryAccentBrush"]
    }
    $badgeText = $Dialog.FindName('TxtAboutBadgeText')
    if ($badgeText -and $Window.Resources.Contains("BodyTextBrush")) {
      $badgeText.Foreground = $Window.Resources["BodyTextBrush"]
    }
    foreach ($name in @('TxtAboutTitle')) {
      $el = $Dialog.FindName($name)
      if ($el -and $Window.Resources.Contains("PrimaryTextBrush")) {
        $el.Foreground = $Window.Resources["PrimaryTextBrush"]
      }
    }
    foreach ($name in @('TxtAboutVersion','TxtAboutAuthor')) {
      $el = $Dialog.FindName($name)
      if ($el -and $Window.Resources.Contains("SecondaryTextBrush")) {
        $el.Foreground = $Window.Resources["SecondaryTextBrush"]
      }
    }
    foreach ($name in @('TxtAboutDescription','TxtAboutNote')) {
      $el = $Dialog.FindName($name)
      if ($el -and $Window.Resources.Contains("BodyTextBrush")) {
        $el.Foreground = $Window.Resources["BodyTextBrush"]
      }
    }
    $btn = $Dialog.FindName('BtnAboutClose')
    if ($btn) {
      if ($Window.Resources.Contains('BodyTextBrush')) { $btn.Foreground = $Window.Resources['BodyTextBrush'] }
      if ($Window.Resources.Contains('ControlBorderBrush')) { $btn.BorderBrush = $Window.Resources['ControlBorderBrush'] }
      if ($Window.Resources.Contains('ControlBackgroundBrush')) { $btn.Background = $Window.Resources['ControlBackgroundBrush'] }
    }
  } catch {}
}

function Update-DarkModeToggleVisual {
  if (-not $TglDarkMode) { return }
  try {
    if ($Script:ActiveTheme -eq "dark") {
      $TglDarkMode.Content = "☀️"
      $TglDarkMode.ToolTip = "Switch to light mode"
    } else {
      $TglDarkMode.Content = "🌙"
      $TglDarkMode.ToolTip = "Switch to dark mode"
    }
  } catch {}
}

function Detect-Tables([ValidateSet("PRIMARY","SECONDARY")]$Platform) {
  # Try to discover SQL backend table names via INFORMATION_SCHEMA
  $schema = Get-Cfg 'SqlSchema' 'dbo'
  $prefix = Get-Cfg 'TablePrefix' 'tb_'
  $force  = [bool](Get-Cfg 'ForceSqlNames' $false)
  if ($force) {
    $map = [ordered]@{
      Tech       = "[{0}].[{1}]" -f $schema,(Get-Cfg 'TechTableName' 'tb_Technician')
      Priv       = "[{0}].[{1}]" -f $schema,(Get-Cfg 'PrivTableName' 'tb_TechPrivilege')
      Verify     = "[{0}].[{1}]" -f $schema,(Get-Cfg 'VerifyTableName' 'tb_TechVerify')
      LabOpsLive = "[{0}].[{1}]" -f $schema,(Get-Cfg 'LabOpsLiveTableName' 'tb_LabOpsLive')
    }
    $Script:TableCache[$Platform] = $map
    return $map
  }
  $c = $null
  try {
    $c = Open-Conn $Platform
    $cmd = $c.CreateCommand()
    $cmd.CommandText = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE IN ('BASE TABLE','VIEW')"
    $adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
    $dt  = New-Object System.Data.DataTable
    [void]$adp.Fill($dt)

    function Pick([string]$pattern) {
      $rows = $dt.Rows | Where-Object { $_.TABLE_NAME -match $pattern }
      $pref = $rows | Where-Object { $_.TABLE_NAME -match '^(tb_|TB_)' }
      if ($pref.Count -gt 0) { $rows = $pref }
      if ($rows.Count -gt 0) {
        return "[{0}].[{1}]" -f $rows[0].TABLE_SCHEMA, $rows[0].TABLE_NAME
      }
      return $null
    }

    $tech = Pick '(^tb_|_)?Technician$'
    $priv = Pick '(^tb_|_)?TechPrivilege$'
    $verf = Pick '(^tb_|_)?TechVerify$'
    $mlive= Pick 'LabOps.*live$|LabOpsLive$|(^tb_|_)?LabOpsLive$'

    if (-not $tech -or -not $priv -or -not $verf) {
      # Hard fallback to dbo.tb_* names seen in Linked Table Manager
      $map = [ordered]@{
        Tech       = "[{0}].[{1}]" -f $schema,(Get-Cfg 'TechTableName' 'tb_Technician')
        Priv       = "[{0}].[{1}]" -f $schema,(Get-Cfg 'PrivTableName' 'tb_TechPrivilege')
        Verify     = "[{0}].[{1}]" -f $schema,(Get-Cfg 'VerifyTableName' 'tb_TechVerify')
        LabOpsLive = if ($mlive) { $mlive } else { "[{0}].[{1}]" -f $schema,(Get-Cfg 'LabOpsLiveTableName' 'tb_LabOpsLive') }
      }
      $Script:TableCache[$Platform] = $map
      return $map
    }

    $map = [ordered]@{
      Tech       = $tech
      Priv       = $priv
      Verify     = $verf
      LabOpsLive = if ($mlive) { $mlive } else { "[{0}].[{1}]" -f $schema,(Get-Cfg 'LabOpsLiveTableName' 'tb_LabOpsLive') }
    }
    $Script:TableCache[$Platform] = $map
    return $map
  }
  finally { Close-Conn $c }
}

function Get-Tables([ValidateSet("PRIMARY","SECONDARY")]$Platform) {
  if ($Script:TableCache.ContainsKey($Platform)) { return $Script:TableCache[$Platform] }
  try {
    return Detect-Tables -Platform $Platform
  } catch {
    # Final hard fallback: use config-driven names
    $schema = Get-Cfg 'SqlSchema' 'dbo'
    $map = [ordered]@{
      Tech       = "[{0}].[{1}]" -f $schema,(Get-Cfg 'TechTableName' 'tb_Technician')
      Priv       = "[{0}].[{1}]" -f $schema,(Get-Cfg 'PrivTableName' 'tb_TechPrivilege')
      Verify     = "[{0}].[{1}]" -f $schema,(Get-Cfg 'VerifyTableName' 'tb_TechVerify')
      LabOpsLive = "[{0}].[{1}]" -f $schema,(Get-Cfg 'LabOpsLiveTableName' 'tb_LabOpsLive')
    }
    $Script:TableCache[$Platform] = $map
    return $map
  }
}

# Diagnostic: show detected tables
function Show-DetectedTables {
  foreach ($plat in @("PRIMARY","SECONDARY")) {
    $t = Get-Tables $plat
    Write-Host "$plat -> Tech=$($t.Tech); Priv=$($t.Priv); Verify=$($t.Verify); LabOpsLive=$($t.LabOpsLive)"
  }
}
function Open-Conn([ValidateSet("PRIMARY","SECONDARY")]$Platform) {
  $dsn = Get-DSN $Platform
  $db  = if ($Platform -eq "PRIMARY") { Get-Cfg 'Database_PRIMARY' '' } else { Get-Cfg 'Database_SECONDARY' '' }
  $dbPart = ""
  if (-not [string]::IsNullOrWhiteSpace($db)) {
    $dbPart = "Database={0};" -f $db
  }
  try {
    $connStr = "DSN={0};{1}Trusted_Connection=Yes;" -f $dsn,$dbPart
    $c = New-Object System.Data.Odbc.OdbcConnection $connStr
    $c.Open();
    $info = if ([string]::IsNullOrWhiteSpace($db)) { "DSN={0}" -f $dsn } else { "DSN={0};DB={1}" -f $dsn,$db }
    Write-Log ("ODBC connected {0} ({1})" -f $info,$Platform) "OK"
    $c
  } catch {
    $info = if ([string]::IsNullOrWhiteSpace($db)) { $dsn } else { "{0};DB={1}" -f $dsn,$db }
    Write-Log "ODBC open failed ($Platform/$info): $($_.Exception.Message)" "ERR"
    throw
  }
}
function Close-Conn($c) { if ($c) { $c.Close(); $c.Dispose() } }

# --------------------------- SOP enforcement ---------------------------
function Get-RolePrivs([ValidateSet("Lab Associate","Certified Lab Tech","CLS","Director")]$Role) {
  switch ($Role) {
    "Lab Associate"      { @("VIEW") }
    "Certified Lab Tech" { @("VIEW","APPROVE","ACCEPT") }
    "CLS"                { @("VIEW","APPROVE","ACCEPT") }
    "Director"           { @("VIEW","APPROVE","ACCEPT","INTERPRETATION","REAPPROVE") }
  }
}
function Enforce-OpsPortal([string]$JobTitle,[bool]$CLSYes,[string[]]$RequestedPrivs) {
  if ($JobTitle -eq "Lab Associate" -and -not $CLSYes) {
    if ($RequestedPrivs | Where-Object { $_ -in @("ACCEPT","APPROVE") }) {
      throw "SOP: Lab Associate without CLS YES cannot receive OpsPortal (ACCEPT/APPROVE)."
    }
  }
}

# --------------------------- Data ops ---------------------------
function Get-Tech([ValidateSet("PRIMARY","SECONDARY")]$Platform,[int]$TechID) {
  $t = Get-Tables $Platform
  $c = Open-Conn $Platform
  try {
    $cmd = $c.CreateCommand()
    $cmd.CommandText = "SELECT T.TechID,T.UserName,T.Area,T.Name,T.Password,T.Multi,
                               V.FirstName,V.MiddleName,V.LastName,V.emplID
                        FROM $($t.Tech) T
                        LEFT JOIN $($t.Verify) V ON T.TechID=V.TechID
                        WHERE T.TechID=?"
    $p=$cmd.CreateParameter(); $p.Value=$TechID; [void]$cmd.Parameters.Add($p)
    $adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
    $dt  = New-Object System.Data.DataTable
    [void]$adp.Fill($dt)

    $pcmd = $c.CreateCommand()
    $pcmd.CommandText = "SELECT Privilege FROM $($t.Priv) WHERE TechID=?"
    $pp=$pcmd.CreateParameter(); $pp.Value=$TechID; [void]$pcmd.Parameters.Add($pp)
    $padp = New-Object System.Data.Odbc.OdbcDataAdapter $pcmd
    $pdt  = New-Object System.Data.DataTable
    [void]$padp.Fill($pdt)

    return @{ Info=$dt; Privs=($pdt.Rows | ForEach-Object {$_.Privilege}) }
  } finally { Close-Conn $c }
}

function Search-Techs([ValidateSet("PRIMARY","SECONDARY")]$Platform,[string]$Text,[int]$Max=$Cfg.MaxRows,[string]$HasPrivilege) {
  $t = Get-Tables $Platform
  $c = Open-Conn $Platform
  try {
    if ($Max -le 0) { $Max = $Cfg.MaxRows }
    # Build WHERE and parameters dynamically to avoid undeclared @q (ODBC uses ? placeholders)
    $where = @()
    $params = New-Object System.Collections.Generic.List[System.Object]

    $trimText = if ($null -eq $Text) { "" } else { $Text.Trim() }
    $q = if ([string]::IsNullOrWhiteSpace($trimText)) { $null } else { "%$trimText%" }
    if ($null -ne $q) {
      $where += "(T.UserName LIKE ? OR T.Name LIKE ? OR V.FirstName LIKE ? OR V.LastName LIKE ? OR CAST(T.TechID AS VARCHAR(20)) LIKE ? OR CAST(V.emplID AS VARCHAR(20)) LIKE ?)"
      1..6 | ForEach-Object { [void]$params.Add($q) }
    }
    $numericSearch = $null
    if (-not [string]::IsNullOrWhiteSpace($trimText)) {
      $tmpInt = 0
      if ([int]::TryParse($trimText, [ref]$tmpInt)) {
        $numericSearch = $tmpInt
      }
    }
    if ($numericSearch -ne $null) {
      $where += "(T.TechID = ? OR V.emplID = ?)"
      [void]$params.Add($numericSearch)
      [void]$params.Add($numericSearch)
    }

    if (-not [string]::IsNullOrWhiteSpace($HasPrivilege)) {
      $where += "EXISTS (SELECT 1 FROM $($t.Priv) P WHERE P.TechID=T.TechID AND P.Privilege=?)"
      [void]$params.Add($HasPrivilege)
    }

    $whereSql = if ($where.Count -gt 0) { "WHERE " + ($where -join " AND ") } else { "" }
    $orderClause = "ORDER BY T.UserName"
    if ([string]::IsNullOrWhiteSpace($trimText) -and [string]::IsNullOrWhiteSpace($HasPrivilege)) {
      # When listing everything, show newest TechIDs first so TOP $Max does not trim recent users
      $orderClause = "ORDER BY T.TechID DESC, T.UserName"
    }

    $sql = @"
SELECT TOP $Max T.TechID,T.UserName,T.Area,T.Name,T.Password,T.Multi,
       V.FirstName,V.MiddleName,V.LastName,V.emplID
FROM $($t.Tech) T
LEFT JOIN $($t.Verify) V ON T.TechID=V.TechID
$whereSql
$orderClause
"@

    $cmd = $c.CreateCommand()
    $cmd.CommandText = $sql
    foreach ($v in $params) {
      $p = $cmd.CreateParameter()
      $p.Value = $v
      [void]$cmd.Parameters.Add($p)
    }

    $adp = New-Object System.Data.Odbc.OdbcDataAdapter $cmd
    $dt  = New-Object System.Data.DataTable
    [void]$adp.Fill($dt)

    # Fetch privileges for each row
    if (-not $dt.Columns.Contains("Privileges")) {
      [void]$dt.Columns.Add("Privileges",[string])
    }
    foreach ($r in $dt.Rows) {
      $pcmd = $c.CreateCommand()
      $pcmd.CommandText = "SELECT Privilege FROM $($t.Priv) WHERE TechID=?"
      $pp=$pcmd.CreateParameter(); $pp.Value=$r.TechID; [void]$pcmd.Parameters.Add($pp)
      $padp = New-Object System.Data.Odbc.OdbcDataAdapter $pcmd
      $pdt  = New-Object System.Data.DataTable
      [void]$padp.Fill($pdt)
      $r["Privileges"] = ($pdt.Rows | ForEach-Object { $_.Privilege }) -join ", "
    }

    return ,$dt
  }
  finally {
    Close-Conn $c
  }
}

function Upsert-Tech {
  param(
    [ValidateSet("PRIMARY","SECONDARY")]$Platform,
    [int]$TechID,[string]$UserName,[string]$Area,[string]$FullName,[string]$Password,
    [string]$FirstName,[string]$MiddleName,[string]$LastName,[int]$EmplID,
    [string[]]$Privileges,[string]$JobTitle,[bool]$CLSYes,
    [switch]$CreateLabOpsLive,[ValidateSet("nonlicensed","licensed")]$LabOpsRole="nonlicensed",
    [ValidateSet("auto","create","update")]$Mode="auto"
  )
  # Validate
  foreach ($p in $Privileges) { if ($p -notin $Cfg.AllowedPrivileges) { throw "Privilege '$p' not allowed." } }
  Enforce-OpsPortal -JobTitle $JobTitle -CLSYes $CLSYes -RequestedPrivs $Privileges

  $t = Get-Tables $Platform
  $c = Open-Conn $Platform
  $tx = $c.BeginTransaction()
  try {
    # exists?
    $ex = $c.CreateCommand(); $ex.Transaction = $tx
    $ex.CommandText = "SELECT COUNT(1) FROM $($t.Tech) WHERE TechID=?"
    $pex=$ex.CreateParameter(); $pex.Value=$TechID; [void]$ex.Parameters.Add($pex)
    $exists = [int]$ex.ExecuteScalar()

    switch ($Mode) {
      "create" { if ($exists -ne 0) { throw "TechID $TechID already exists. Use Update instead." } }
      "update" { if ($exists -eq 0) { throw "TechID $TechID not found. Use New User instead." } }
      default {}
    }

    if ($exists -eq 0) {
      $cmd = $c.CreateCommand(); $cmd.Transaction=$tx
      $cmd.CommandText = "INSERT INTO $($t.Tech) (TechID,UserName,Area,Password,Name,Multi) VALUES (?,?,?,?,?,?)"
      foreach ($v in @($TechID,$UserName,$Area,$Password,$FullName,"FALSE")) { $p=$cmd.CreateParameter(); $p.Value=$v; [void]$cmd.Parameters.Add($p) }
      [void]$cmd.ExecuteNonQuery()
    } else {
      $cmd = $c.CreateCommand(); $cmd.Transaction=$tx
      $cmd.CommandText = "UPDATE $($t.Tech) SET UserName=?,Area=?,Password=?,Name=? WHERE TechID=?"
      foreach ($v in @($UserName,$Area,$Password,$FullName,$TechID)) { $p=$cmd.CreateParameter(); $p.Value=$v; [void]$cmd.Parameters.Add($p) }
      [void]$cmd.ExecuteNonQuery()
    }

    # verify upsert
    $vex = $c.CreateCommand(); $vex.Transaction=$tx
    $vex.CommandText = "SELECT COUNT(1) FROM $($t.Verify) WHERE TechID=?"
    $pv=$vex.CreateParameter(); $pv.Value=$TechID; [void]$vex.Parameters.Add($pv)
    if ([int]$vex.ExecuteScalar() -eq 0) {
      $vc = $c.CreateCommand(); $vc.Transaction=$tx
      $vc.CommandText = "INSERT INTO $($t.Verify) (TechID,UserName,Area,Name,FirstName,MiddleName,LastName,emplID) VALUES (?,?,?,?,?,?,?,?)"
      foreach ($v in @($TechID,$UserName,$Area,$FullName,$FirstName,$MiddleName,$LastName,$EmplID)) { $p=$vc.CreateParameter(); $p.Value=$v; [void]$vc.Parameters.Add($p) }
      [void]$vc.ExecuteNonQuery()
    } else {
      $vc = $c.CreateCommand(); $vc.Transaction=$tx
      $vc.CommandText = "UPDATE $($t.Verify) SET UserName=?,Area=?,Name=?,FirstName=?,MiddleName=?,LastName=?,emplID=? WHERE TechID=?"
      foreach ($v in @($UserName,$Area,$FullName,$FirstName,$MiddleName,$LastName,$EmplID,$TechID)) { $p=$vc.CreateParameter(); $p.Value=$v; [void]$vc.Parameters.Add($p) }
      [void]$vc.ExecuteNonQuery()
    }

    # reset + set privileges
    $del = $c.CreateCommand(); $del.Transaction=$tx
    $del.CommandText = "DELETE FROM $($t.Priv) WHERE TechID=?"
    $pd=$del.CreateParameter(); $pd.Value=$TechID; [void]$del.Parameters.Add($pd)
    [void]$del.ExecuteNonQuery()
    foreach ($p in ($Privileges | Select-Object -Unique)) {
      $ins = $c.CreateCommand(); $ins.Transaction=$tx
      $ins.CommandText = "INSERT INTO $($t.Priv) (TechID,Privilege) VALUES (?,?)"
      foreach ($v in @($TechID,$p)) { $pi=$ins.CreateParameter(); $pi.Value=$v; [void]$ins.Parameters.Add($pi) }
      [void]$ins.ExecuteNonQuery()
    }

    if ($CreateLabOpsLive) {
      $ml = $c.CreateCommand(); $ml.Transaction=$tx
      $ml.CommandText = "SELECT COUNT(1) FROM $($t.LabOpsLive) WHERE TechID=?"
      $pm=$ml.CreateParameter(); $pm.Value=$TechID; [void]$ml.Parameters.Add($pm)
      if ([int]$ml.ExecuteScalar() -eq 0) {
        $mi = $c.CreateCommand(); $mi.Transaction=$tx
        $mi.CommandText = "INSERT INTO $($t.LabOpsLive) (TechID,Role) VALUES (?,?)"
        foreach ($v in @($TechID,$LabOpsRole)) { $pmi=$mi.CreateParameter(); $pmi.Value=$v; [void]$mi.Parameters.Add($pmi) }
        [void]$mi.ExecuteNonQuery()
      }
    }

    if ($Cfg.ReadOnly) { throw "ReadOnly is enabled – transaction preview only (no commit)." }
    $tx.Commit()
    Audit -Action "upsert" -Details @{ platform=$Platform; techid=$TechID; uname=$UserName; privs=$Privileges } -Result "success"
    return "OK"
  } catch {
    try { $tx.Rollback() | Out-Null } catch {}
    Audit -Action "upsert" -Details @{ platform=$Platform; techid=$TechID } -Result "error" -Message $_
    throw
  } finally { Close-Conn $c }
}

function Disable-Tech([ValidateSet("PRIMARY","SECONDARY")]$Platform,[int]$TechID) {
  $t = Get-Tables $Platform
  $c = Open-Conn $Platform
  $tx= $c.BeginTransaction()
  try {
    $d = $c.CreateCommand(); $d.Transaction=$tx
    $d.CommandText = "DELETE FROM $($t.Priv) WHERE TechID=?"
    $pd=$d.CreateParameter(); $pd.Value=$TechID; [void]$d.Parameters.Add($pd)
    [void]$d.ExecuteNonQuery()

    $u = $c.CreateCommand(); $u.Transaction=$tx
    $u.CommandText = "UPDATE $($t.Tech) SET Area='INFORMATION', Name='TERMINATED' WHERE TechID=?"
    $pu=$u.CreateParameter(); $pu.Value=$TechID; [void]$u.Parameters.Add($pu)
    [void]$u.ExecuteNonQuery()

    # Also blank Verify name fields (Option B)
    $uv = $c.CreateCommand(); $uv.Transaction=$tx
    $uv.CommandText = "UPDATE $($t.Verify) SET Name='', FirstName='', MiddleName='', LastName='' WHERE TechID=?"
    $puv=$uv.CreateParameter(); $puv.Value=$TechID; [void]$uv.Parameters.Add($puv)
    [void]$uv.ExecuteNonQuery()

    if ($Cfg.ReadOnly) { throw "ReadOnly is enabled – transaction preview only (no commit)." }
    $tx.Commit()
    Audit -Action "disable" -Details @{ platform=$Platform; techid=$TechID } -Result "success"
    "OK"
  } catch {
    try { $tx.Rollback() | Out-Null } catch {}
    Audit -Action "disable" -Details @{ platform=$Platform; techid=$TechID } -Result "error" -Message $_
    throw
  } finally { Close-Conn $c }
}

function Mirror-User([ValidateSet("PRIMARY","SECONDARY")]$Platform,[int]$SourceTechID,[int]$TargetTechID,[string]$TargetUserName,[string]$TargetFullName,[string]$TargetArea,[int]$TargetEmplID) {
  $src = Get-Tech -Platform $Platform -TechID $SourceTechID
  if (-not $src.Info.Rows) { throw "Source TechID not found." }
  $privs = $src.Privs
  $fn="";$mn="";$ln=""
  if ($TargetFullName -and $TargetFullName.Contains(",")) { $parts=$TargetFullName.Split(","); $ln=$parts[0].Trim(); $fn=$parts[1].Trim() }
  Upsert-Tech -Platform $Platform -TechID $TargetTechID -UserName $TargetUserName -Area $TargetArea -FullName $TargetFullName -Password $TargetEmplID `
    -FirstName $fn -MiddleName $mn -LastName $ln -EmplID $TargetEmplID `
    -Privileges $privs -JobTitle "Certified Lab Tech" -CLSYes:$true -CreateLabOpsLive:$true -LabOpsRole "nonlicensed"
}

function Resolve-TechIdValue([object]$Value,[string]$Label) {
  if ($null -eq $Value) { throw "TechID $Label required." }
  if ($Value -is [int]) { return [int]$Value }
  if ($Value -is [long]) { return [int]$Value }
  if ($Value -is [double]) { return [int][math]::Round($Value,0) }
  if ($Value -is [string]) {
    $trim = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($trim)) { throw "TechID $Label required." }
    $num = 0
    if (-not [int]::TryParse($trim, [ref]$num)) { throw "TechID $Label must be numeric." }
    return $num
  }
  if ($Value -is [System.Collections.Hashtable]) {
    foreach ($key in @('TechID','Value','Text','Content')) {
      if ($Value.ContainsKey($key) -and $Value[$key]) {
        return Resolve-TechIdValue $Value[$key] $Label
      }
    }
  }
  if ($Value -is [System.Management.Automation.PSCustomObject]) {
    foreach ($prop in @('TechID','Value','Text','Content')) {
      $match = $Value.PSObject.Properties.Match($prop)
      if ($match -and $match[0].Value) {
        return Resolve-TechIdValue $match[0].Value $Label
      }
    }
  }
  try { return [int]$Value } catch {}
  throw "TechID $Label must be numeric."
}

function Compare-Users([ValidateSet("PRIMARY","SECONDARY")]$Platform,$A,$B) {
  $idA = Resolve-TechIdValue $A "A"
  $idB = Resolve-TechIdValue $B "B"

  $a = Get-Tech -Platform $Platform -TechID $idA
  $b = Get-Tech -Platform $Platform -TechID $idB

  if (-not $a.Info -or $a.Info.Rows.Count -eq 0) { throw "TechID $idA not found on $Platform." }
  if (-not $b.Info -or $b.Info.Rows.Count -eq 0) { throw "TechID $idB not found on $Platform." }

  $rowA = $a.Info.Rows[0]
  $rowB = $b.Info.Rows[0]

  $getVal = {
    param($row,$column)
    if (-not $row) { return "" }
    if (-not $row.Table.Columns.Contains($column)) { return "" }
    $val = $row[$column]
    if ($null -eq $val -or $val -is [System.DBNull]) { return "" }
    return $val
  }

  $aUserName   = &$getVal $rowA "UserName"
  $aArea       = &$getVal $rowA "Area"
  $aFullName   = &$getVal $rowA "Name"
  $aFirstName  = &$getVal $rowA "FirstName"
  $aMiddleName = &$getVal $rowA "MiddleName"
  $aLastName   = &$getVal $rowA "LastName"
  $aEmplID     = &$getVal $rowA "emplID"

  $bUserName   = &$getVal $rowB "UserName"
  $bArea       = &$getVal $rowB "Area"
  $bFullName   = &$getVal $rowB "Name"
  $bFirstName  = &$getVal $rowB "FirstName"
  $bMiddleName = &$getVal $rowB "MiddleName"
  $bLastName   = &$getVal $rowB "LastName"
  $bEmplID     = &$getVal $rowB "emplID"

  $privA = @()
  if ($a.Privs) { $privA = @($a.Privs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }
  $privB = @()
  if ($b.Privs) { $privB = @($b.Privs | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }

  $missingInB = @()
  if ($privA.Count -gt 0) {
    $missingInB = @($privA | Where-Object { $_ -notin $privB })
  }
  $missingInA = @()
  if ($privB.Count -gt 0) {
    $missingInA = @($privB | Where-Object { $_ -notin $privA })
  }

  [pscustomobject]@{
    Platform = $Platform
    A = [pscustomobject]@{
      TechID    = $idA
      UserName  = $aUserName
      Area      = $aArea
      FullName  = $aFullName
      FirstName = $aFirstName
      MiddleName= $aMiddleName
      LastName  = $aLastName
      EmplID    = $aEmplID
      Privileges= $privA
    }
    B = [pscustomobject]@{
      TechID    = $idB
      UserName  = $bUserName
      Area      = $bArea
      FullName  = $bFullName
      FirstName = $bFirstName
      MiddleName= $bMiddleName
      LastName  = $bLastName
      EmplID    = $bEmplID
      Privileges= $privB
    }
    MissingInB = $missingInB
    MissingInA = $missingInA
  }
}

function Format-CompareRows($cmp) {
  if (-not $cmp) { return @() }
  $rows = New-Object System.Collections.Generic.List[object]

  $rows.Add([pscustomobject]@{
    Field="TechID"; A=$cmp.A.TechID; B=$cmp.B.TechID;
    Difference = $(if ($cmp.A.TechID -eq $cmp.B.TechID) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="User Name"; A=$cmp.A.UserName; B=$cmp.B.UserName;
    Difference = $(if ($cmp.A.UserName -eq $cmp.B.UserName) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="Area"; A=$cmp.A.Area; B=$cmp.B.Area;
    Difference = $(if ($cmp.A.Area -eq $cmp.B.Area) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="Full Name"; A=$cmp.A.FullName; B=$cmp.B.FullName;
    Difference = $(if ($cmp.A.FullName -eq $cmp.B.FullName) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="First Name"; A=$cmp.A.FirstName; B=$cmp.B.FirstName;
    Difference = $(if ($cmp.A.FirstName -eq $cmp.B.FirstName) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="Middle Name"; A=$cmp.A.MiddleName; B=$cmp.B.MiddleName;
    Difference = $(if ($cmp.A.MiddleName -eq $cmp.B.MiddleName) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="Last Name"; A=$cmp.A.LastName; B=$cmp.B.LastName;
    Difference = $(if ($cmp.A.LastName -eq $cmp.B.LastName) { "Match" } else { "Mismatch" })
  })
  $rows.Add([pscustomobject]@{
    Field="Employee ID"; A=$cmp.A.EmplID; B=$cmp.B.EmplID;
    Difference = $(if ($cmp.A.EmplID -eq $cmp.B.EmplID) { "Match" } else { "Mismatch" })
  })

  $privA = ""
  if ($cmp.A.Privileges -and $cmp.A.Privileges.Count -gt 0) {
    $privA = ($cmp.A.Privileges -join ", ")
  }
  $privB = ""
  if ($cmp.B.Privileges -and $cmp.B.Privileges.Count -gt 0) {
    $privB = ($cmp.B.Privileges -join ", ")
  }
  $privDiffParts = @()
  if ($cmp.MissingInB -and $cmp.MissingInB.Count -gt 0) {
    $privDiffParts += ("Missing in B: " + ($cmp.MissingInB -join ", "))
  }
  if ($cmp.MissingInA -and $cmp.MissingInA.Count -gt 0) {
    $privDiffParts += ("Missing in A: " + ($cmp.MissingInA -join ", "))
  }
  if ($privDiffParts.Count -eq 0) { $privDiffParts = @("Match") }

  $rows.Add([pscustomobject]@{
    Field="Privileges"; A=$privA; B=$privB;
    Difference=($privDiffParts -join " | ")
  })

  return $rows
}

function Show-AboutDialog {
  $aboutXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="About $AppName"
        Height="320" Width="460"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        WindowStyle="ToolWindow"
        Background="White"
        ShowInTaskbar="False">
  <Grid x:Name="AboutRoot" Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <StackPanel Orientation="Horizontal">
      <Grid Width="52" Height="52">
        <Ellipse x:Name="AboutBadgeShape" Fill="#1C3F80"/>
        <TextBlock x:Name="TxtAboutBadgeText" Text="AM" Foreground="White" FontWeight="Bold" FontSize="20" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Grid>
      <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
        <TextBlock x:Name="TxtAboutTitle" Text="$AppName" FontSize="20" FontWeight="Bold"/>
        <TextBlock x:Name="TxtAboutVersion" Text="Version $AppVer" FontSize="12"/>
      </StackPanel>
    </StackPanel>
    <TextBlock x:Name="TxtAboutAuthor" Grid.Row="1" Text="Author: Andy Sendouw" Margin="0,16,0,0" FontSize="14"/>
    <TextBlock x:Name="TxtAboutDescription" Grid.Row="2" Text="Access Manager streamlines LabOps/OpsPortal account changes with guardrails for SOP compliance and auditing." TextWrapping="Wrap" Margin="0,16,0,0"/>
    <TextBlock x:Name="TxtAboutNote" Grid.Row="3" Text="For support or feedback, contact the Access Manager team." TextWrapping="Wrap" Margin="0,12,0,0"/>
    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,24,0,0">
      <Button x:Name="BtnAboutClose" Content="Close" Width="80" Padding="10,4"/>
    </StackPanel>
  </Grid>
</Window>
"@
  $reader = New-Object System.Xml.XmlNodeReader ([xml]$aboutXaml)
  $dlg = [Windows.Markup.XamlReader]::Load($reader)
  $btnClose = $dlg.FindName('BtnAboutClose')
  if ($btnClose) {
    $btnClose.Add_Click({ param($sender,$args) $dlg.Close() })
  }
  if ($Window) {
    $dlg.Owner = $Window
  }
  try { Apply-ThemeToAbout $dlg } catch {}
  $dlg.ShowDialog() | Out-Null
}

# --------------------------- XAML UI ---------------------------
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$AppName v$AppVer"
        Height="780" Width="1280" MinHeight="720" MinWidth="1100"
        WindowStartupLocation="CenterScreen"
        SizeToContent="Manual"
        SnapsToDevicePixels="True" UseLayoutRounding="True"
        Background="{DynamicResource WindowBackgroundBrush}">
  <Window.Resources>
    <SolidColorBrush x:Key="WindowBackgroundBrush" Color="#FAFBFF"/>
    <SolidColorBrush x:Key="SurfaceBrush" Color="#F3F6FB"/>
    <SolidColorBrush x:Key="PanelBorderBrush" Color="#D6DEED"/>
    <SolidColorBrush x:Key="PrimaryAccentBrush" Color="#1C3F80"/>
    <SolidColorBrush x:Key="PrimaryTextBrush" Color="#1C3F80"/>
    <SolidColorBrush x:Key="SecondaryTextBrush" Color="#4C5466"/>
    <SolidColorBrush x:Key="BodyTextBrush" Color="#2C3140"/>
    <SolidColorBrush x:Key="TipTextBrush" Color="#6C7285"/>
    <SolidColorBrush x:Key="GroupBackgroundBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="GridRowBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="GridAltRowBrush" Color="#EDF2FC"/>
    <SolidColorBrush x:Key="GridCompareAltRowBrush" Color="#F4F7FD"/>
    <SolidColorBrush x:Key="GridHeaderBackgroundBrush" Color="#E4EAF7"/>
    <SolidColorBrush x:Key="GridHeaderForegroundBrush" Color="#1C3F80"/>
    <SolidColorBrush x:Key="ControlBackgroundBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="ControlBorderBrush" Color="#D6DEED"/>
    <Style x:Key="FormLabelStyle" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,2"/>
      <Setter Property="Foreground" Value="{DynamicResource BodyTextBrush}"/>
    </Style>
    <Style x:Key="FormInputStyle" TargetType="TextBox">
      <Setter Property="Margin" Value="0,0,0,6"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Foreground" Value="{DynamicResource BodyTextBrush}"/>
      <Setter Property="Background" Value="{DynamicResource ControlBackgroundBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource ControlBorderBrush}"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="10,8,10,10"/>
      <Setter Property="BorderBrush" Value="{DynamicResource PanelBorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Margin" Value="0,12,0,0"/>
      <Setter Property="Background" Value="{DynamicResource GroupBackgroundBrush}"/>
      <Setter Property="Foreground" Value="{DynamicResource BodyTextBrush}"/>
      <Setter Property="HeaderTemplate">
        <Setter.Value>
          <DataTemplate>
            <TextBlock Text="{Binding}" FontWeight="SemiBold" Margin="0,0,0,2" Foreground="{DynamicResource PrimaryTextBrush}"/>
          </DataTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DataGridHeaderStyle" TargetType="{x:Type DataGridColumnHeader}">
      <Setter Property="Foreground" Value="{DynamicResource GridHeaderForegroundBrush}"/>
      <Setter Property="Background" Value="{DynamicResource GridHeaderBackgroundBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource PanelBorderBrush}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="6,4"/>
    </Style>
  </Window.Resources>
  <DockPanel LastChildFill="True" Margin="10" TextElement.Foreground="{DynamicResource BodyTextBrush}">

    <!-- Top toolbar -->
    <Border DockPanel.Dock="Top" Background="{StaticResource SurfaceBrush}" CornerRadius="8" Padding="12" Margin="0,0,0,8" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1">
      <Grid VerticalAlignment="Center">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="16"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Grid.Column="0">
          <Grid Width="40" Height="40">
            <Ellipse Fill="{DynamicResource PrimaryAccentBrush}"/>
            <TextBlock Text="AM" Foreground="White" FontWeight="Bold" FontSize="16" HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Grid>
          <StackPanel Margin="10,0,0,0" VerticalAlignment="Center">
            <TextBlock Text="AccessManager" FontWeight="Bold" FontSize="16" Foreground="{DynamicResource PrimaryTextBrush}"/>
            <TextBlock Text="Author: Andy Sendouw" Foreground="{DynamicResource SecondaryTextBrush}" FontSize="12"/>
          </StackPanel>
          <Button x:Name="BtnAbout" Content="ℹ" FontSize="16" Width="32" Height="32" Margin="16,0,0,0" Padding="0" VerticalAlignment="Center" ToolTip="About AccessManager"/>
          <ToggleButton x:Name="TglDarkMode" Content="🌙" FontSize="16" Width="32" Height="32" Margin="8,0,0,0" VerticalAlignment="Center" ToolTip="Toggle dark mode" HorizontalContentAlignment="Center" VerticalContentAlignment="Center"/>
        </StackPanel>

        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Grid.Column="2">
          <TextBlock Text="Platform:" Margin="0,0,6,0" VerticalAlignment="Center"/>
          <ComboBox x:Name="CboPlatform" Width="90" SelectedIndex="1" VerticalAlignment="Center" SelectedValuePath="Content"
                    Foreground="{DynamicResource BodyTextBrush}" Background="{DynamicResource ControlBackgroundBrush}"
                    BorderBrush="{DynamicResource ControlBorderBrush}">
            <ComboBoxItem Content="PRIMARY"/>
            <ComboBoxItem Content="SECONDARY"/>
          </ComboBox>
          <TextBlock x:Name="TxtPlatformStatus" Margin="8,0,0,0" VerticalAlignment="Center" FontWeight="SemiBold" Foreground="{DynamicResource PrimaryTextBrush}"/>
          <CheckBox x:Name="ChkRO" Content="Read-Only (safe)" Margin="12,0,0,0" IsChecked="True" VerticalAlignment="Center"
                    Foreground="{DynamicResource BodyTextBrush}" Background="Transparent" BorderBrush="{DynamicResource ControlBorderBrush}"/>
          <TextBox x:Name="TxtSearch" Width="360" Margin="12,0,0,0" VerticalAlignment="Center" Foreground="{DynamicResource BodyTextBrush}" Background="{DynamicResource ControlBackgroundBrush}" BorderBrush="{DynamicResource ControlBorderBrush}"/>
          <Button x:Name="BtnSearch" Content="Search / List" Margin="8,0,0,0" Width="120" VerticalAlignment="Center" Padding="12,4"/>
          <TextBlock Text="Has Privilege:" Margin="16,0,6,0" VerticalAlignment="Center"/>
          <ComboBox x:Name="CboHasPriv" Width="160" VerticalAlignment="Center" Foreground="{DynamicResource BodyTextBrush}"
                    Background="{DynamicResource ControlBackgroundBrush}" BorderBrush="{DynamicResource ControlBorderBrush}">
            <ComboBoxItem Content=""/>
            <ComboBoxItem Content="VIEW"/>
            <ComboBoxItem Content="ACCEPT"/>
            <ComboBoxItem Content="APPROVE"/>
            <ComboBoxItem Content="REAPPROVE"/>
            <ComboBoxItem Content="INTERPRETATION"/>
            <ComboBoxItem Content="LIMITA"/>
            <ComboBoxItem Content="LIMITB"/>
          </ComboBox>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Main content -->
    <ScrollViewer DockPanel.Dock="Top" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
      <Grid Margin="0,0,0,0">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Results grid -->
        <DataGrid x:Name="GridUsers" Grid.Row="0" Grid.ColumnSpan="3" Margin="0,0,0,12"
                  Height="320" MaxHeight="320"
                  AutoGenerateColumns="True" IsReadOnly="True"
                  EnableRowVirtualization="True" EnableColumnVirtualization="True"
                  ColumnWidth="*" RowHeaderWidth="0" HeadersVisibility="Column"
                  ColumnHeaderStyle="{DynamicResource DataGridHeaderStyle}"
                  Foreground="{DynamicResource BodyTextBrush}"
                  Background="{DynamicResource SurfaceBrush}"
                  ScrollViewer.CanContentScroll="True"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
                  AlternationCount="2"
                  RowBackground="{DynamicResource GridRowBrush}"
                  AlternatingRowBackground="{DynamicResource GridAltRowBrush}"
                  GridLinesVisibility="Horizontal"
                  BorderBrush="{StaticResource PanelBorderBrush}"
                  BorderThickness="1"/>

        <!-- Editor area -->
        <Border Grid.Row="1" Background="{DynamicResource GroupBackgroundBrush}" CornerRadius="8" Padding="12"
                BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1">
          <Grid Margin="0,0,0,0">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>

          <!-- Row 0 -->
          <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,12,6">
            <StackPanel Margin="0,0,0,6">
              <TextBlock Text="TechID" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtTechID" MinWidth="100" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
            <StackPanel Margin="0,0,0,6">
              <TextBlock Text="UserName" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtUser" MinWidth="160" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
            <StackPanel>
              <TextBlock Text="Area" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtArea" Text="labops" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
          </StackPanel>

          <StackPanel Grid.Row="0" Grid.Column="1" Margin="0,0,12,6">
            <StackPanel Margin="0,0,0,6">
              <TextBlock Text="Full Name (Last, First)" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtFull" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
            <StackPanel>
              <TextBlock Text="Password (emplID#)" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtPass" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
          </StackPanel>

          <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,6">
            <StackPanel Margin="0,0,0,6">
              <TextBlock Text="Job Title" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtJob" Text="Certified Lab Tech" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <TextBlock Text="CLS YES?" Margin="0,0,8,0" VerticalAlignment="Center"/>
              <CheckBox x:Name="ChkCLS" VerticalAlignment="Center"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
            </StackPanel>
          </StackPanel>

          <!-- Row 1 -->
          <StackPanel Grid.Row="1" Grid.Column="0" Orientation="Horizontal" Margin="0,0,12,6">
            <StackPanel Margin="0,0,12,0" MinWidth="150">
              <TextBlock Text="First Name" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtFirst" MinWidth="140" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
            <StackPanel MinWidth="120">
              <TextBlock Text="Middle Name" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtMiddle" MinWidth="120" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
          </StackPanel>

          <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Horizontal" Margin="0,0,12,6">
            <StackPanel Margin="0,0,12,0" MinWidth="190">
              <TextBlock Text="Last Name" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtLast" MinWidth="170" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
            <StackPanel MinWidth="150">
              <TextBlock Text="Employee ID" Style="{StaticResource FormLabelStyle}"/>
              <TextBox x:Name="TxtEmplID" MinWidth="130" HorizontalAlignment="Stretch" Style="{StaticResource FormInputStyle}"/>
            </StackPanel>
          </StackPanel>

          <GroupBox Header="Privileges" Grid.Row="1" Grid.Column="2" Margin="0,0,0,6" Padding="10,6">
            <WrapPanel x:Name="WrapPrivs" Margin="0">
              <CheckBox Content="VIEW" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
              <CheckBox Content="ACCEPT" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
              <CheckBox Content="APPROVE" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
              <CheckBox Content="REAPPROVE" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
              <CheckBox Content="INTERPRETATION" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
              <CheckBox Content="LIMITA" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
              <CheckBox Content="LIMITB" Margin="0,0,6,6"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="Transparent"
                        BorderBrush="{DynamicResource ControlBorderBrush}"/>
            </WrapPanel>
          </GroupBox>

          <!-- Row 2 -->
          <StackPanel Grid.Row="2" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,0,12,0">
            <Button x:Name="BtnCreate" Content="New User" Width="130" Margin="0,0,8,0" Padding="12,4"/>
            <Button x:Name="BtnUpdate" Content="Update" Width="120" Margin="0,0,8,0" Padding="12,4"/>
            <Button x:Name="BtnDisable" Content="Disable User" Width="130" Margin="0,0,8,0" Padding="12,4"/>
            <Button x:Name="BtnMirror" Content="Mirror From…" Width="130" Margin="0,0,8,0" Padding="12,4"/>
          </StackPanel>

          <TextBlock Grid.Row="2" Grid.Column="2" Text="Tip: First run is Read-Only. Uncheck to commit writes."
                     VerticalAlignment="Center" HorizontalAlignment="Right" Foreground="{DynamicResource TipTextBrush}" FontStyle="Italic"/>

          <GroupBox Grid.Row="3" Grid.ColumnSpan="3" Header="Compare TechIDs" Margin="0,12,0,0">
            <Grid Margin="0,6,0,0">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="TechID A:" VerticalAlignment="Center"/>
                <TextBox x:Name="TxtCompareA" Width="100" Margin="6,0,12,0" VerticalAlignment="Center" Style="{StaticResource FormInputStyle}"/>
                <TextBlock Text="TechID B:" VerticalAlignment="Center"/>
                <TextBox x:Name="TxtCompareB" Width="100" Margin="6,0,12,0" VerticalAlignment="Center" Style="{StaticResource FormInputStyle}"/>
                <Button x:Name="BtnCompareRun" Content="Compare" Width="110" Margin="0,0,12,0" VerticalAlignment="Center" Padding="10,4"/>
                <TextBlock x:Name="TxtCompareStatus" VerticalAlignment="Center" Foreground="{DynamicResource SecondaryTextBrush}" Margin="8,0,0,0"/>
              </StackPanel>
              <DataGrid x:Name="GridCompare" Grid.Row="1" Margin="0,8,0,0" Height="180"
                        AutoGenerateColumns="False" IsReadOnly="True"
                        ColumnWidth="*" RowHeaderWidth="0" HeadersVisibility="Column"
                        ColumnHeaderStyle="{DynamicResource DataGridHeaderStyle}"
                        Foreground="{DynamicResource BodyTextBrush}"
                        Background="{DynamicResource SurfaceBrush}"
                        CanUserAddRows="False" CanUserDeleteRows="False"
                        ScrollViewer.CanContentScroll="True"
                        ScrollViewer.VerticalScrollBarVisibility="Auto"
                        AlternationCount="2"
                        RowBackground="{DynamicResource GridRowBrush}"
                        AlternatingRowBackground="{DynamicResource GridCompareAltRowBrush}"
                        BorderBrush="{StaticResource PanelBorderBrush}"
                        BorderThickness="1">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Field" Binding="{Binding Field}" Width="160"/>
                  <DataGridTextColumn Header="A" Binding="{Binding A}" Width="*"/>
                  <DataGridTextColumn Header="B" Binding="{Binding B}" Width="*"/>
                  <DataGridTextColumn Header="Difference" Binding="{Binding Difference}" Width="220"/>
                </DataGrid.Columns>
              </DataGrid>
            </Grid>
          </GroupBox>
          </Grid>
        </Border>
      </Grid>
    </ScrollViewer>
  </DockPanel>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$Xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# --------------------------- Grab controls ---------------------------
$CboPlatform = $Window.FindName('CboPlatform')
$ChkRO       = $Window.FindName('ChkRO')
$TxtSearch   = $Window.FindName('TxtSearch')
$BtnSearch   = $Window.FindName('BtnSearch')
$CboHasPriv  = $Window.FindName('CboHasPriv')
$TxtPlatformStatus = $Window.FindName('TxtPlatformStatus')
$GridUsers   = $Window.FindName('GridUsers')

$BtnAbout     = $Window.FindName('BtnAbout')
$TglDarkMode  = $Window.FindName('TglDarkMode')

$TxtTechID   = $Window.FindName('TxtTechID')
$TxtUser     = $Window.FindName('TxtUser')
$TxtArea     = $Window.FindName('TxtArea')
$TxtFull     = $Window.FindName('TxtFull')
$TxtPass     = $Window.FindName('TxtPass')
$TxtJob      = $Window.FindName('TxtJob')
$ChkCLS      = $Window.FindName('ChkCLS')

$TxtFirst    = $Window.FindName('TxtFirst')
$TxtMiddle   = $Window.FindName('TxtMiddle')
$TxtLast     = $Window.FindName('TxtLast')
$TxtEmplID   = $Window.FindName('TxtEmplID')

$WrapPrivs   = $Window.FindName('WrapPrivs')

$BtnCreate        = $Window.FindName('BtnCreate')
$BtnUpdate        = $Window.FindName('BtnUpdate')
$BtnDisable       = $Window.FindName('BtnDisable')
$BtnMirror        = $Window.FindName('BtnMirror')
$TxtCompareA      = $Window.FindName('TxtCompareA')
$TxtCompareB      = $Window.FindName('TxtCompareB')
$BtnCompareRun    = $Window.FindName('BtnCompareRun')
$GridCompare      = $Window.FindName('GridCompare')
$TxtCompareStatus = $Window.FindName('TxtCompareStatus')

# --------------------------- Helpers for UI ---------------------------
function Set-ActivePlatform([string]$Platform) {
  if ([string]::IsNullOrWhiteSpace($Platform)) { return }
  $Script:ActivePlatform = $Platform
}
function Current-Platform {
  if (-not [string]::IsNullOrWhiteSpace($Script:ActivePlatform)) { return $Script:ActivePlatform }
  if (-not [string]::IsNullOrWhiteSpace($Cfg.DefaultPlatform)) { return $Cfg.DefaultPlatform }
  return $null
}
function Update-PlatformIndicator {
  param(
    [Nullable[int]]$RowCount = $null,
    [string]$Context = $null
  )
  $plat = Current-Platform
  if (-not $plat) {
    if ($TxtPlatformStatus) { $TxtPlatformStatus.Text = "Active platform not selected" }
    $Window.Title = "$AppName v$AppVer"
    return
  }
  $dsn = ""
  try { $dsn = Get-DSN $plat } catch { $dsn = "" }
  $db = ""
  try {
    $db = if ($plat -eq "PRIMARY") { Get-Cfg 'Database_PRIMARY' '' } else { Get-Cfg 'Database_SECONDARY' '' }
  } catch { $db = "" }
  $statusParts = New-Object System.Collections.Generic.List[string]
  if ([string]::IsNullOrWhiteSpace($dsn)) {
    [void]$statusParts.Add(("Active: {0}" -f $plat))
  } else {
    if ([string]::IsNullOrWhiteSpace($db)) {
      [void]$statusParts.Add(("Active: {0} (DSN {1})" -f $plat,$dsn))
    } else {
      [void]$statusParts.Add(("Active: {0} (DSN {1} • DB {2})" -f $plat,$dsn,$db))
    }
  }
  if ($RowCount -ne $null) {
    [void]$statusParts.Add(("Rows: {0}" -f $RowCount))
  }
  if (-not [string]::IsNullOrWhiteSpace($Context)) {
    [void]$statusParts.Add($Context)
  }
  if ($TxtPlatformStatus) {
    $TxtPlatformStatus.Text = ($statusParts -join " • ")
  }
  $Window.Title = "{0} v{1} [{2}]" -f $AppName,$AppVer,$plat
  $shouldPersist = $false
  if ($Cfg.DefaultPlatform -ne $plat) {
    $Cfg.DefaultPlatform = $plat
    $shouldPersist = $true
  }
  if ($shouldPersist) {
    Save-Cfg $Cfg
  }
}
function Handle-PlatformChange([string]$PlatformOverride) {
  if (-not [string]::IsNullOrWhiteSpace($PlatformOverride)) {
    Set-ActivePlatform $PlatformOverride
  }
  $plat = Current-Platform
  if (-not $plat) { return }
  $dsn = ""
  try { $dsn = Get-DSN $plat } catch { $dsn = "" }
  if ([string]::IsNullOrWhiteSpace($dsn)) {
    Write-Log ("Platform switched to {0}" -f $plat)
  } else {
    Write-Log ("Platform switched to {0} dsn={1}" -f $plat,$dsn)
  }
  Update-PlatformIndicator -RowCount $null -Context ("Switched {0}" -f (Get-Date -Format 'HH:mm:ss'))
  $hasSearch = -not [string]::IsNullOrWhiteSpace($TxtSearch.Text)
  $hasFilter = $CboHasPriv -and $CboHasPriv.SelectedIndex -gt 0
  if ($hasSearch -or $hasFilter) {
    try { Refresh-Grid } catch { Write-Log ("Refresh on platform change failed: {0}" -f $_.Exception.Message) "ERR" }
  } else {
    if ($GridUsers) { $GridUsers.ItemsSource = $null }
    Update-PlatformIndicator -RowCount 0 -Context "Results cleared"
  }
  if ($GridCompare) { $GridCompare.ItemsSource = $null }
  if ($TxtCompareStatus) { $TxtCompareStatus.Text = "" }
}
function Get-CheckedPrivs {
  $list = @()
  foreach ($child in $WrapPrivs.Children) { if ($child.IsChecked) { $list += $child.Content.ToString() } }
  $list
}
function Set-CheckedPrivs([string[]]$privs) {
  foreach ($child in $WrapPrivs.Children) {
    $child.IsChecked = ($privs -contains $child.Content.ToString())
  }
}

function Invoke-TechSave {
  param(
    [ValidateSet("create","update")]$Mode,
    [string]$SuccessMessage="Saved."
  )
  $Cfg.ReadOnly = [bool]$ChkRO.IsChecked; Save-Cfg $Cfg
  $privs = Get-CheckedPrivs
  $plat  = Current-Platform

  $parsedTechID = 0
  if (-not [int]::TryParse($TxtTechID.Text, [ref]$parsedTechID)) {
    throw "TechID required (numeric TechID)."
  }
  $parsedEmplID = 0
  if (-not [int]::TryParse($TxtEmplID.Text, [ref]$parsedEmplID)) {
    throw "emplID required (numeric emplID)."
  }

  $res = Upsert-Tech -Platform $plat `
    -TechID $parsedTechID `
    -UserName $TxtUser.Text -Area $TxtArea.Text -FullName $TxtFull.Text -Password $TxtPass.Text `
    -FirstName $TxtFirst.Text -MiddleName $TxtMiddle.Text -LastName $TxtLast.Text -EmplID $parsedEmplID `
    -Privileges $privs -JobTitle $TxtJob.Text -CLSYes:([bool]$ChkCLS.IsChecked) `
    -Mode $Mode

  if ($res -eq "OK") {
    $actionLabel = if ($SuccessMessage) { $SuccessMessage.Trim() } else { "" }
    $actionLabel = $actionLabel.TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($actionLabel)) {
      $actionLabel = if ($Mode -eq "create") { "Created" } else { "Updated" }
    }
    $dsn = ""
    try { $dsn = Get-DSN $plat } catch { $dsn = "" }
    $locationLabel = if ([string]::IsNullOrWhiteSpace($dsn)) { $plat } else { "{0} (DSN {1})" -f $plat,$dsn }
    $dialogMessage = "{0} on {1}." -f $actionLabel,$locationLabel
    [System.Windows.MessageBox]::Show($dialogMessage,"OK") | Out-Null
    $Script:LastActionNote = "{0} TechID {1} on {2}" -f $actionLabel,$parsedTechID,$plat
    if ([string]::IsNullOrWhiteSpace($dsn)) {
      Write-Log ("{0} TechID {1} on {2}" -f $actionLabel,$parsedTechID,$plat)
    } else {
      Write-Log ("{0} TechID {1} on {2} dsn={3}" -f $actionLabel,$parsedTechID,$plat,$dsn)
    }
  }
  Refresh-Grid
}

# Normalizes various data types for DataGrid ItemsSource
function To-ItemsSource($obj) {
  if ($null -eq $obj) { return $null }
  $t = $obj.GetType().FullName
  switch ($t) {
    'System.Data.DataTable' { return $obj.DefaultView }
    'System.Data.DataView'  { return $obj }
    'System.Data.DataSet'   { if ($obj.Tables.Count -gt 0) { return $obj.Tables[0].DefaultView } else { return $null } }
    default {
      if ($obj -is [System.Collections.IEnumerable]) { return $obj }
      else { return $null }
    }
  }
}

function Refresh-Grid {
  try {
    $p = Current-Platform
    $searchText = if ($TxtSearch.Text) { $TxtSearch.Text.Trim() } else { "" }

    # Safely read selected privilege (can be null)
    $hp = ""
    try {
      if ($CboHasPriv.SelectedItem -and $CboHasPriv.SelectedItem.Content) {
        $hp = [string]$CboHasPriv.SelectedItem.Content
      }
    } catch { $hp = "" }

    # Convert to $null when empty to match Search-Techs expectation
    $hpArg = $null
    if (-not [string]::IsNullOrWhiteSpace($hp)) { $hpArg = $hp }

    $dt = Search-Techs -Platform $p -Text $TxtSearch.Text -HasPrivilege $hpArg
    $rowCount = 0
    if ($dt -and ($dt -is [System.Data.DataTable])) { $rowCount = $dt.Rows.Count }

    $src = To-ItemsSource $dt
    if ($null -eq $src) {
      throw ("Unexpected result type from Search-Techs: {0}" -f ($dt.GetType().FullName))
    }
    $GridUsers.ItemsSource = $src
    $stamp = Get-Date -Format 'HH:mm:ss'
    $dsn = ""
    try { $dsn = Get-DSN $p } catch { $dsn = "" }
    $db = ""
    try { $db = if ($p -eq "PRIMARY") { Get-Cfg 'Database_PRIMARY' '' } else { Get-Cfg 'Database_SECONDARY' '' } } catch { $db = "" }
    $contextParts = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($Script:LastActionNote)) {
      [void]$contextParts.Add($Script:LastActionNote)
      $Script:LastActionNote = $null
    }
    [void]$contextParts.Add(("Last search {0}" -f $stamp))
    $context = ($contextParts -join "; ")
    Update-PlatformIndicator -RowCount $rowCount -Context $context
    $privFilter = if ($hpArg) { $hpArg } else { "" }
    $logMsg = "Search refreshed platform={0} rows={1} text='{2}' priv='{3}'"
    $logArgs = @($p,$rowCount,$searchText,$privFilter)
    if (-not [string]::IsNullOrWhiteSpace($dsn)) {
      if ([string]::IsNullOrWhiteSpace($db)) {
        $logMsg = "Search refreshed platform={0} dsn={1} rows={2} text='{3}' priv='{4}'"
        $logArgs = @($p,$dsn,$rowCount,$searchText,$privFilter)
      } else {
        $logMsg = "Search refreshed platform={0} dsn={1} db={2} rows={3} text='{4}' priv='{5}'"
        $logArgs = @($p,$dsn,$db,$rowCount,$searchText,$privFilter)
      }
    }
    Write-Log ([string]::Format($logMsg, $logArgs))
  }
  catch {
    [System.Windows.MessageBox]::Show(("Search failed: {0}" -f $_.Exception.Message),"Error","OK","Error") | Out-Null
  }
}

# Set initial values from config
if (-not [string]::IsNullOrWhiteSpace($Cfg.DefaultPlatform)) {
  try { $CboPlatform.SelectedValue = $Cfg.DefaultPlatform } catch {}
  Set-ActivePlatform $Cfg.DefaultPlatform
}
$preferredTheme = (Get-Cfg 'Theme' 'light')
Apply-ThemeByName $preferredTheme
if ($TglDarkMode) {
  try { $TglDarkMode.IsChecked = ($preferredTheme -eq 'dark') } catch {}
  Update-DarkModeToggleVisual
}
$ChkRO.IsChecked = $Cfg.ReadOnly
Update-PlatformIndicator

if ($BtnAbout) {
  $BtnAbout.Add_Click({
    try { Show-AboutDialog } catch {
      [System.Windows.MessageBox]::Show(("Unable to open About window: {0}" -f $_.Exception.Message),"About","OK","Error") | Out-Null
    }
  })
}
if ($TglDarkMode) {
  $TglDarkMode.Add_Checked({
    if ($Script:ActiveTheme -ne "dark") {
      Apply-ThemeByName "dark" -Persist
      Update-DarkModeToggleVisual
    }
  })
  $TglDarkMode.Add_Unchecked({
    if ($Script:ActiveTheme -ne "light") {
      Apply-ThemeByName "light" -Persist
      Update-DarkModeToggleVisual
    }
  })
}

$CboPlatform.Add_SelectionChanged({ param($sender,$eventArgs)
  try {
    $selPlat = $null
    try {
      if ($sender -and $sender.SelectedValue) { $selPlat = $sender.SelectedValue.ToString().Trim() }
    } catch {}
    if ([string]::IsNullOrWhiteSpace($selPlat)) {
      try {
        if ($sender -and $sender.SelectedItem -and $sender.SelectedItem.Content) {
          $selPlat = $sender.SelectedItem.Content.ToString().Trim()
        }
      } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($selPlat) -and $sender -and $sender.Text) {
      $selPlat = $sender.Text.Trim()
    }
    Handle-PlatformChange -PlatformOverride $selPlat
  } catch {
    Write-Log ("Platform change handler failed: {0}" -f $_.Exception.Message) "ERR"
  }
})

$BtnSearch.Add_Click({ Refresh-Grid })

$GridUsers.Add_SelectionChanged({
  try {
    if ($GridUsers.SelectedItem) {
      $r = $GridUsers.SelectedItem.Row
      $TxtTechID.Text = $r.TechID
      $TxtUser.Text   = $r.UserName
      $TxtArea.Text   = $r.Area
      $TxtFull.Text   = $r.Name
      $TxtPass.Text   = $r.Password
      $TxtFirst.Text  = $r.FirstName
      $TxtMiddle.Text = $r.MiddleName
      $TxtLast.Text   = $r.LastName
      $TxtEmplID.Text = $r.emplID
      Set-CheckedPrivs (($r.Privileges -as [string]) -split ",\s*")
    }
  } catch {}
})

$BtnCreate.Add_Click({
  try {
    Invoke-TechSave -Mode "create" -SuccessMessage "Created."
  } catch {
    [System.Windows.MessageBox]::Show("Create failed: $($_.Exception.Message)","Error","OK","Error") | Out-Null
  }
})

$BtnUpdate.Add_Click({
  try {
    Invoke-TechSave -Mode "update" -SuccessMessage "Updated."
  } catch {
    [System.Windows.MessageBox]::Show("Update failed: $($_.Exception.Message)","Error","OK","Error") | Out-Null
  }
})

$BtnDisable.Add_Click({
  try {
    $Cfg.ReadOnly = [bool]$ChkRO.IsChecked; Save-Cfg $Cfg
    $tmpTechID = 0
    if (-not [int]::TryParse($TxtTechID.Text, [ref]$tmpTechID)) { throw "TechID required (numeric TechID)." }
    $plat = Current-Platform
    $null = Disable-Tech -Platform $plat -TechID ([int]$TxtTechID.Text)
    [System.Windows.MessageBox]::Show("Disabled: privileges cleared, Area=INFORMATION, Name=TERMINATED. TechVerify name fields cleared.","OK") | Out-Null
    Refresh-Grid
  } catch {
    [System.Windows.MessageBox]::Show("Disable failed: $($_.Exception.Message)","Error","OK","Error") | Out-Null
  }
})

$BtnMirror.Add_Click({
  try {
    $Cfg.ReadOnly = [bool]$ChkRO.IsChecked; Save-Cfg $Cfg
    $src = [Microsoft.VisualBasic.Interaction]::InputBox("Source TechID to mirror from:","","")
    if ([string]::IsNullOrWhiteSpace($src)) { return }
    $plat = Current-Platform
    $null = Mirror-User -Platform $plat -SourceTechID ([int]$src) -TargetTechID ([int]$TxtTechID.Text) `
      -TargetUserName $TxtUser.Text -TargetFullName $TxtFull.Text -TargetArea $TxtArea.Text -TargetEmplID ([int]$TxtEmplID.Text)
    [System.Windows.MessageBox]::Show("Mirrored privileges from $src.","OK") | Out-Null
    Refresh-Grid
  } catch {
    [System.Windows.MessageBox]::Show("Mirror failed: $($_.Exception.Message)","Error","OK","Error") | Out-Null
  }
})

$BtnCompareRun.Add_Click({
  try {
    $plat = Current-Platform
    if ([string]::IsNullOrWhiteSpace($plat)) { throw "Select a platform first." }

    $aText = if ($TxtCompareA.Text) { $TxtCompareA.Text.Trim() } else { "" }
    $bText = if ($TxtCompareB.Text) { $TxtCompareB.Text.Trim() } else { "" }
    if ([string]::IsNullOrWhiteSpace($aText) -or [string]::IsNullOrWhiteSpace($bText)) {
      throw "Enter TechID values for both A and B."
    }

    $aID = 0
    if (-not [int]::TryParse($aText, [ref]$aID)) { throw "TechID A must be numeric." }
    $bID = 0
    if (-not [int]::TryParse($bText, [ref]$bID)) { throw "TechID B must be numeric." }

    $cmp = Compare-Users -Platform $plat -A $aID -B $bID
    $rows = Format-CompareRows $cmp
    $GridCompare.ItemsSource = $rows
    $stamp = Get-Date -Format 'HH:mm:ss'
    if ($TxtCompareStatus) {
      $TxtCompareStatus.Text = "Compared $aID ↔ $bID at $stamp"
    }
    $Script:LastActionNote = "Compared $aID↔$bID on $plat"
    Write-Log ("Compared TechID {0} with {1} on {2}" -f $aID,$bID,$plat)
  } catch {
    if ($GridCompare) { $GridCompare.ItemsSource = $null }
    if ($TxtCompareStatus) { $TxtCompareStatus.Text = "" }
    [System.Windows.MessageBox]::Show(("Compare failed: {0}" -f $_.Exception.Message),"Compare","OK","Error") | Out-Null
  }
})

# On close, persist platform + read-only
$Window.Add_Closing({
  try {
    $platClose = Current-Platform
    if (-not [string]::IsNullOrWhiteSpace($platClose)) {
      $Cfg.DefaultPlatform = $platClose
    }
    $Cfg.ReadOnly = [bool]$ChkRO.IsChecked
    Save-Cfg $Cfg
  } catch {}
})


# Run
$app = New-Object System.Windows.Application
$app.Run($Window) | Out-Null
