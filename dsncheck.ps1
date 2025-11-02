# ===========================
# Legacy ODBC / TLS Helper (Win11)
# Interactive menu, run as Administrator
# ===========================

# --- Elevation check ---
$curr = [Security.Principal.WindowsIdentity]::GetCurrent()
$prn  = New-Object Security.Principal.WindowsPrincipal($curr)
if (-not $prn.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Host "Please run PowerShell as Administrator." -ForegroundColor Red
  break
}

$ErrorActionPreference = 'Stop'
$LogsDir = Join-Path $env:ProgramData "Win7-Schannel-Compat"
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null
$RunId   = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile = Join-Path $LogsDir "Run-$RunId.log"

function Log($m,[string]$lvl='INFO'){
  ("{0} [{1}] {2}" -f (Get-Date -Format s),$lvl,$m) | Tee-Object -FilePath $LogFile -Append | Out-Null
}

function Backup-Reg {
  param([Parameter(Mandatory)][string]$key,[Parameter(Mandatory)][string]$name)
  try {
    if (Test-Path "Registry::$key") {
      $dst = Join-Path $LogsDir "$name-$RunId.reg"
      & reg.exe export $key $dst /y | Out-Null
      Log "Backup OK: $key -> $dst"
    } else {
      Log "Backup skipped (key not found): $key" 'WARN'
    }
  } catch {
    Log "Backup error for $key : $_" 'WARN'
  }
}

function Ensure-PolicyPaths {
  foreach ($p in @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL',
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Cryptography\Configuration\SSL'
  )) { if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null } }
}

# --- Win7-compat Schannel posture ---
function Set-Win7SchannelCompat {
  param([switch]$AlsoWinNode,[switch]$AllowNullExport)

  Log "Applying Win7-compat Schannel posture..."

  $polA = 'HKLM:\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002'
  $polB = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Cryptography\Configuration\SSL\00010002'
  $targets = @($polA) + $(if($AlsoWinNode){@($polB)})

  $cipherOrder = @'
TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA384_P256,TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA384_P384,TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA256_P256,TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA256_P384,TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA_P256,TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA_P384,TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA_P256,TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA_P384,TLS_DHE_RSA_WITH_AES_256_GCM_SHA384,TLS_DHE_RSA_WITH_AES_128_GCM_SHA256,TLS_DHE_RSA_WITH_AES_256_CBC_SHA,TLS_DHE_RSA_WITH_AES_128_CBC_SHA,TLS_RSA_WITH_AES_256_GCM_SHA384,TLS_RSA_WITH_AES_128_GCM_SHA256,TLS_RSA_WITH_AES_256_CBC_SHA256,TLS_RSA_WITH_AES_128_CBC_SHA256,TLS_RSA_WITH_AES_256_CBC_SHA,TLS_RSA_WITH_AES_128_CBC_SHA,TLS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384_P384,TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256_P256,TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256_P384,TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384_P384,TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256_P256,TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256_P384,TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA_P256,TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA_P384,TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA_P256,TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA_P384,TLS_DHE_DSS_WITH_AES_256_CBC_SHA256,TLS_DHE_DSS_WITH_AES_128_CBC_SHA256,TLS_DHE_DSS_WITH_AES_256_CBC_SHA,TLS_DHE_DSS_WITH_AES_128_CBC_SHA,TLS_RSA_WITH_3DES_EDE_CBC_SHA,TLS_DHE_DSS_WITH_3DES_EDE_CBC_SHA,TLS_RSA_WITH_RC4_128_SHA,TLS_RSA_WITH_RC4_128_MD5,TLS_RSA_WITH_NULL_SHA256,TLS_RSA_WITH_NULL_SHA,SSL_CK_RC4_128_WITH_MD5,SSL_CK_DES_192_EDE3_CBC_WITH_MD5
'@.Trim()

  foreach($p in $targets){
    if(-not (Test-Path $p)){ New-Item -Path $p -Force | Out-Null }
    if(Get-ItemProperty -Path $p -Name 'Functions' -ErrorAction SilentlyContinue){
      Remove-ItemProperty -Path $p -Name 'Functions' -ErrorAction SilentlyContinue
    }
    New-ItemProperty -Path $p -Name 'Functions' -PropertyType String -Value ($cipherOrder -replace '\s*,\s*',',') -Force | Out-Null
    $count = ((Get-ItemProperty -Path $p -Name 'Functions').Functions -split ',').Count
    Log "Cipher suite order applied at $p (REG_SZ; $count entries)"
  }

  $pb = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'
  foreach($role in 'Client','Server'){
    foreach($proto in 'TLS 1.3'){ $k=Join-Path $pb (Join-Path $proto $role); if(-not(Test-Path $k)){New-Item -Path $k -Force|Out-Null}; New-ItemProperty -Path $k -Name Enabled -Type DWord -Value 0 -Force|Out-Null; New-ItemProperty -Path $k -Name DisabledByDefault -Type DWord -Value 1 -Force|Out-Null }
    foreach($proto in 'TLS 1.0','TLS 1.1','TLS 1.2'){ $k=Join-Path $pb (Join-Path $proto $role); if(-not(Test-Path $k)){New-Item -Path $k -Force|Out-Null}; New-ItemProperty -Path $k -Name Enabled -Type DWord -Value 1 -Force|Out-Null; New-ItemProperty -Path $k -Name DisabledByDefault -Type DWord -Value 0 -Force|Out-Null }
    # Uncomment to allow SSL 3.0 if truly required:
    # $k=Join-Path $pb (Join-Path 'SSL 3.0' $role); if(-not(Test-Path $k)){New-Item -Path $k -Force|Out-Null}; New-ItemProperty -Path $k -Name Enabled -Type DWord -Value 1 -Force|Out-Null; New-ItemProperty -Path $k -Name DisabledByDefault -Type DWord -Value 0 -Force|Out-Null
  }

  $cb = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers'
  function Set-C($n,$v){ $k=Join-Path $cb $n; if(-not(Test-Path $k)){New-Item -Path $k -Force|Out-Null}; New-ItemProperty -Path $k -Name Enabled -Type DWord -Value $v -Force | Out-Null; Log "Cipher $n -> $v" }
  Set-C 'AES 128/128' 0xFFFFFFFF
  Set-C 'AES 256/256' 0xFFFFFFFF
  Set-C 'Triple DES 168' 0xFFFFFFFF
  Set-C 'RC4 128/128' 0xFFFFFFFF
  Set-C 'RC4 64/128'  0xFFFFFFFF
  Set-C 'RC4 56/128'  0xFFFFFFFF
  Set-C 'RC4 40/128'  0xFFFFFFFF

  foreach($n in 'NULL','DES 56/56','DES 40/56','EXPORT1024','EXPORT56','EXPORT40'){
    Set-C $n 0  # OFF by default; use Option 6 to enable if absolutely required
  }

  # WinHTTP defaults (TLS1.0/1.1/1.2)
  $roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'
  )
  $mask = 0x200 -bor 0x400 -bor 0x800
  foreach($r in $roots){ if(-not(Test-Path $r)){New-Item -Path $r -Force|Out-Null}; New-ItemProperty -Path $r -Name DefaultSecureProtocols -Type DWord -Value $mask -Force | Out-Null }
  Log "Win7-compat Schannel posture applied. Reboot required."
}

# --- ODBC helpers ---
function Get-ODBCMap {
  [pscustomobject]@{
    User32_Path    = 'HKCU:\SOFTWARE\ODBC\ODBC.INI'
    User32_Sources = 'HKCU:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources'
    Sys32_Path     = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI'
    Sys32_Sources  = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC Data Sources'
    Sys32_Drivers  = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI'
    User64_Path    = 'HKCU:\SOFTWARE\ODBC\ODBC.INI'
    User64_Sources = 'HKCU:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources'
    Sys64_Path     = 'HKLM:\SOFTWARE\ODBC\ODBC.INI'
    Sys64_Sources  = 'HKLM:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources'
    Sys64_Drivers  = 'HKLM:\SOFTWARE\ODBC\ODBCINST.INI'
  }
}

# Robust DSN lister (finds Sys32 reliably)
function List-DSNs {
  $m = Get-ODBCMap
  $rows = @()

  function Get-ScopeDSNs($scope) {
    $iniPath = $m."$scope`_Path"
    $srcPath = $m."$scope`_Sources"
    $map = @{}

    if (Test-Path $srcPath) {
      $props = (Get-ItemProperty -Path $srcPath).PSObject.Properties |
               Where-Object { $_.Name -notmatch '^PS' }
      foreach ($p in $props) { $map[$p.Name] = "" + $p.Value }
    }
    if (Test-Path $iniPath) {
      Get-ChildItem -Path $iniPath -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.PSChildName -ne 'ODBC Data Sources') {
          if (-not $map.ContainsKey($_.PSChildName)) { $map[$_.PSChildName] = $null }
        }
      }
    }
    foreach ($name in $map.Keys) {
      $drv = $map[$name]
      $dsnKey = Join-Path $iniPath $name
      if (-not $drv -and (Test-Path $dsnKey)) {
        try {
          $p = Get-ItemProperty -Path $dsnKey -ErrorAction SilentlyContinue
          if ($p.Driver) { $drv = $p.Driver }
        } catch {}
      }
      [pscustomobject]@{ Scope=$scope; Name=$name; Driver=$drv }
    }
  }

  foreach ($scope in 'Sys32','User32','Sys64','User64') { $rows += Get-ScopeDSNs $scope }
  $rows | Sort-Object Scope, Name
}

function Remove-DSNKey {
  param([ValidateSet('User32','Sys32','User64','Sys64')]$Scope,[string]$Name)
  $m=Get-ODBCMap; $ini=$m."$Scope`_Path"; $src=$m."$Scope`_Sources"
  $k = Join-Path $ini $Name
  if(Test-Path $k){ Remove-Item -Path $k -Recurse -Force }
  if(Test-Path $src){ Remove-ItemProperty -Path $src -Name $Name -ErrorAction SilentlyContinue }
}

function Ensure-32bitSystemDSN {
  param([string]$Name,[string]$DriverName,[hashtable]$Attributes)
  $m=Get-ODBCMap; $dsnRoot=$m.Sys32_Path; $src=$m.Sys32_Sources; $drvRoot=$m.Sys32_Drivers
  if(-not(Test-Path (Join-Path $drvRoot $DriverName))){ throw "32-bit driver '$DriverName' not found under $drvRoot" }
  foreach($p in @($dsnRoot,$src)){ if(-not(Test-Path $p)){ New-Item -Path $p -Force|Out-Null } }
  $dsnKey = Join-Path $dsnRoot $Name
  if(-not(Test-Path $dsnKey)){ New-Item -Path $dsnKey -Force | Out-Null }
  foreach ($k in $Attributes.Keys){ New-ItemProperty -Path $dsnKey -Name $k -Value $Attributes[$k] -Force | Out-Null }
  New-ItemProperty -Path $src -Name $Name -Value $DriverName -Force | Out-Null
  Write-Host "32-bit System DSN '$Name' created/updated (Driver=$DriverName)." -ForegroundColor Green
}

# --- NEW: Clean DSNs, keep only "default" (name you specify) ---
function Remove-AllDSNsExcept {
  param(
    [Parameter(Mandatory)][ValidateSet('Sys32','User32','Sys64','User64','All')]$Scope,
    [Parameter()][string]$KeepName
  )
  $targets = if ($Scope -eq 'All') { @('Sys32','User32','Sys64','User64') } else { @($Scope) }
  $m = Get-ODBCMap
  $removed = 0

  foreach ($sc in $targets) {
    $iniPath = $m."$sc`_Path"
    $srcPath = $m."$sc`_Sources"
    if (-not (Test-Path $iniPath) -and -not (Test-Path $srcPath)) { continue }

    # Build a full set of names in this scope
    $names = New-Object System.Collections.Generic.HashSet[string]
    if (Test-Path $srcPath) {
      $props = (Get-ItemProperty -Path $srcPath).PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }
      foreach ($p in $props) { $null = $names.Add($p.Name) }
    }
    if (Test-Path $iniPath) {
      Get-ChildItem -Path $iniPath -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.PSChildName -ne 'ODBC Data Sources') { $null = $names.Add($_.PSChildName) }
      }
    }

    foreach ($n in $names) {
      if ([string]::IsNullOrWhiteSpace($KeepName) -or $n -ne $KeepName) {
        Remove-DSNKey -Scope $sc -Name $n
        Log "Removed DSN '$n' from $sc"
        $removed++
      }
    }
  }

  Write-Host ("Removed {0} DSN(s). {1}" -f $removed, $(if($KeepName){"Kept '$KeepName'."}else{"Kept none."})) -ForegroundColor Yellow
  return $removed
}

# --- Menu + actions ---
function Show-Menu {
  Clear-Host
  Write-Host "=========== Legacy ODBC / TLS Helper ===========" -ForegroundColor Yellow
  Write-Host "1) Apply Win7-compatible TLS/SChannel posture"
  Write-Host "2) Create/Update 32-bit System DSN (and clean dupes)"
  Write-Host "3) List current DSNs"
  Write-Host "4) CLEAN DSNs (wipe; keep only default)"
  Write-Host "5) ENABLE NULL/EXPORT ciphers (DANGEROUS)"
  Write-Host "6) Exit"
  Write-Host "Logs: $LogFile"
  Write-Host "================================================"
  Read-Host "Choose 1-6"
}

function Do-ApplySchannel {
  Ensure-PolicyPaths
  # Backups (skip if absent)
  Backup-Reg 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL' 'Policy-Crypto-SSL'
  Backup-Reg 'HKLM\SOFTWARE\Policies\Microsoft\Windows\Cryptography\Configuration\SSL' 'Policy-Win-Crypto-SSL'
  Backup-Reg 'HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL' 'Schannel-Tree'
  Backup-Reg 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings' 'WinINET'
  Backup-Reg 'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings' 'WinINET-WoW'

  $also = (Read-Host "Also write cipher order to Windows\... policy path? (Y/N) [N]")
  $also = if ([string]::IsNullOrWhiteSpace($also)) {'N'} else {$also}
  Set-Win7SchannelCompat -AlsoWinNode:($also -match '^(?i)y$')

  Write-Host "`nApplied. **Reboot required** for TLS changes to take effect." -ForegroundColor Cyan
  Read-Host "Press Enter to continue..."
}

function Do-EnableNullExport {
  Write-Host "`nThis enables NULL/EXPORT/weak DES ciphers. Use ONLY if the peer strictly requires it." -ForegroundColor Red
  $go = (Read-Host "Type I-UNDERSTAND to proceed, anything else to cancel")
  if ($go -notmatch '^(?i)I-UNDERSTAND$') { return }
  $cb = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Ciphers'
  foreach($n in 'NULL','DES 56/56','DES 40/56','EXPORT1024','EXPORT56','EXPORT40'){
    $k=Join-Path $cb $n
    if(-not(Test-Path $k)){ New-Item -Path $k -Force | Out-Null }
    New-ItemProperty -Path $k -Name 'Enabled' -Type DWord -Value 0xFFFFFFFF -Force | Out-Null
    Log "Cipher $n -> ENABLED (legacy/unsafe)"
  }
  Write-Host "Enabled legacy ciphers. **Reboot required**." -ForegroundColor Yellow
  Read-Host "Press Enter to continue..."
}

function Do-CreateDSN {
  $name = Read-Host "DSN Name"
  $drv  = Read-Host "Driver Name (e.g. 'SQL Server' or 'SQL Server Native Client 11.0')"
  $srv  = Read-Host "Server (e.g. SQLHOST\INSTANCE)"
  $db   = Read-Host "Database"
  $tc   = (Read-Host "Trusted_Connection (Yes/No) [Yes]"); if ([string]::IsNullOrWhiteSpace($tc)) { $tc='Yes' }
  $enc  = (Read-Host "Encrypt (Yes/No) [No]");              if ([string]::IsNullOrWhiteSpace($enc)) { $enc='No' }
  $tsc  = (Read-Host "TrustServerCertificate (Yes/No) [Yes]"); if ([string]::IsNullOrWhiteSpace($tsc)) { $tsc='Yes' }

  # Clean duplicates of same DSN from wrong hives
  foreach($scope in 'User64','Sys64','User32'){ Remove-DSNKey -Scope $scope -Name $name }

  $attrs = @{
    Server=$srv; Database=$db;
    Trusted_Connection=$tc; Encrypt=$enc; TrustServerCertificate=$tsc
  }
  try {
    Ensure-32bitSystemDSN -Name $name -DriverName $drv -Attributes $attrs
  } catch {
    Write-Host "Error creating DSN: $($_.Exception.Message)" -ForegroundColor Red
    Log "Ensure-32bitSystemDSN error: $_" 'WARN'
  }
  Read-Host "Press Enter to continue..."
}

function Do-CleanDSN {
  Write-Host "`n--- CLEAN DSNs ---" -ForegroundColor Yellow
  Write-Host "Scopes: 1=Sys32 (System 32-bit), 2=User32, 3=Sys64, 4=User64, 5=All"
  $s = Read-Host "Choose scope (1-5) [1]"
  switch ($s) {
    '2' { $scope = 'User32' }
    '3' { $scope = 'Sys64' }
    '4' { $scope = 'User64' }
    '5' { $scope = 'All' }
    default { $scope = 'Sys32' }
  }
  Write-Host "`nCurrent DSNs before cleanup:" -ForegroundColor Cyan
  List-DSNs | Format-Table -Auto

  $keep = Read-Host "Type the DSN name to KEEP (leave blank to keep none)"
  $confirm = Read-Host "This will DELETE all other DSNs in $scope. Type DELETE to confirm"
  if ($confirm -ne 'DELETE') { Write-Host "Canceled." -ForegroundColor Yellow; return }

  $count = Remove-AllDSNsExcept -Scope $scope -KeepName $keep
  Write-Host "`nAfter cleanup:" -ForegroundColor Cyan
  List-DSNs | Format-Table -Auto
  Write-Host ""
  Read-Host "Press Enter to continue..."
}

# --- Main loop ---
while ($true) {
  switch (Show-Menu) {
    '1' { Do-ApplySchannel }
    '2' { Do-CreateDSN }
    '3' {
      Write-Host "`nCurrent DSNs (by scope):" -ForegroundColor Cyan
      List-DSNs | Format-Table -Auto
      Write-Host ""
      Read-Host "Press Enter to continue..."
    }
    '4' { Do-CleanDSN }
    '5' { Do-EnableNullExport }
    '6' { break }
    default { Write-Host "Invalid choice." -ForegroundColor Red; Start-Sleep -Milliseconds 800 }
  }
}

Write-Host "Done. If you changed TLS/Schannel, REBOOT this machine." -ForegroundColor Yellow
Write-Host "Log: $LogFile"