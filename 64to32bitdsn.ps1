<# 
.SYNOPSIS
  Clone all 64-bit System DSNs to 32-bit System DSNs.

.DESCRIPTION
  - Reads DSNs from: HKLM:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources
  - Writes DSNs to:  HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources
  - Copies each DSN's key/values and updates "Driver" path heuristically to 32-bit.
  - Maps the ODBC Data Sources driver name to a 32-bit installed driver when available.
  - Uses ShouldProcess so you can dry-run with -WhatIf. Use -Force to overwrite.

.PARAMETER Include
  Only process DSNs whose names match these wildcard patterns (e.g. "Prod*","SalesDB").

.PARAMETER Exclude
  Skip DSNs whose names match these wildcard patterns.

.PARAMETER Force
  Overwrite existing 32-bit DSNs.

.EXAMPLE
  .\Clone-SystemDSN64to32.ps1 -WhatIf
  (Shows what would be done without writing.)

.EXAMPLE
  .\Clone-SystemDSN64to32.ps1 -Force
  (Clones all and overwrites any existing 32-bit DSNs.)
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
  [string[]]$Include,
  [string[]]$Exclude,
  [switch]$Force
)

function Test-Admin {
  $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
  $wp = New-Object Security.Principal.WindowsPrincipal($wi)
  return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
  throw "Please run this script in an elevated PowerShell session (Run as Administrator)."
}

$reg64_DSNList   = 'HKLM:\SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources'
$reg64_DSNRoot   = 'HKLM:\SOFTWARE\ODBC\ODBC.INI'
$reg32_DSNList   = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources'
$reg32_DSNRoot   = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI'

# Installed 32-bit ODBC drivers live here (by friendly name)
$reg32_DriverRoot = 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI'

# Ensure 32-bit DSN roots exist
if (-not (Test-Path $reg32_DSNRoot))   { New-Item -Path $reg32_DSNRoot   -Force | Out-Null }
if (-not (Test-Path $reg32_DSNList))   { New-Item -Path $reg32_DSNList   -Force | Out-Null }

# Helper: wildcard include/exclude filter
function Should-ProcessDSN([string]$Name) {
  if ($Include -and -not ($Include | Where-Object { $Name -like $_ })) { return $false }
  if ($Exclude -and     ($Exclude | Where-Object { $Name -like $_ }))  { return $false }
  return $true
}

# Helper: try to map a 64-bit "Driver" value to a 32-bit path
function Convert-To32BitDriverPath([string]$DriverValue) {
  if ([string]::IsNullOrWhiteSpace($DriverValue)) { return $DriverValue }

  $new = $DriverValue

  # Common path swaps
  $new = $new -replace '\\System32\\', '\\SysWOW64\\'                       # DLLs under system
  $new = $new -replace 'C:\\Program Files\\', 'C:\\Program Files (x86)\\'   # vendor installs

  # If original used %SystemRoot%, keep it (SysWOW64 will be picked above if explicit System32 present)
  return $new
}

# Helper: find a 32-bit installed driver with a given friendly name
function Test-32BitDriverExists([string]$FriendlyName) {
  $driverKey = Join-Path $reg32_DriverRoot $FriendlyName
  return (Test-Path $driverKey)
}

# Read 64-bit DSN list
if (-not (Test-Path $reg64_DSNList)) {
  Write-Warning "No 64-bit System DSNs found at: $reg64_DSNList"
  return
}

$sourceDSNs = Get-ItemProperty -Path $reg64_DSNList |
  Select-Object -ExcludeProperty PSPath,PSParentPath,PSChildName,PSDrive,PSProvider |
  ForEach-Object {
    $_.psobject.Properties |
      Where-Object { $_.MemberType -eq 'NoteProperty' } |
      ForEach-Object {
        [PSCustomObject]@{
          Name       = $_.Name
          DriverName = [string]$_.Value  # this is the "friendly driver name" used in ODBC Data Sources
        }
      }
  }

if (-not $sourceDSNs) {
  Write-Warning "No 64-bit System DSNs enumerated."
  return
}

$results = @()

foreach ($dsn in $sourceDSNs) {
  $dsnName   = $dsn.Name
  $drvName64 = $dsn.DriverName

  if (-not (Should-ProcessDSN $dsnName)) { 
    $results += [PSCustomObject]@{ DSN=$dsnName; Status='Skipped (filter)'; Note='Filtered by Include/Exclude' }
    continue 
  }

  $srcKey = Join-Path $reg64_DSNRoot $dsnName
  $dstKey = Join-Path $reg32_DSNRoot $dsnName

  if (-not (Test-Path $srcKey)) {
    $results += [PSCustomObject]@{ DSN=$dsnName; Status='Skipped (missing)'; Note="Missing source key $srcKey" }
    continue
  }

  $exists32 = Test-Path $dstKey
  if ($exists32 -and -not $Force) {
    $results += [PSCustomObject]@{ DSN=$dsnName; Status='Exists (kept)'; Note='Use -Force to overwrite' }
    continue
  }

  $action = if ($exists32) { "Overwrite 32-bit DSN '$dsnName'" } else { "Create 32-bit DSN '$dsnName'" }
  if ($PSCmdlet.ShouldProcess($dsnName, $action)) {
    # (Re)create destination DSN key
    if ($exists32) {
      Remove-Item -Path $dstKey -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -Path $dstKey -Force | Out-Null

    # Copy all properties from 64-bit DSN to 32-bit DSN
    $props = Get-ItemProperty -Path $srcKey |
      Select-Object -ExcludeProperty PSPath,PSParentPath,PSChildName,PSDrive,PSProvider

    foreach ($p in $props.psobject.Properties) {
      $name  = $p.Name
      $value = $p.Value

      # Heuristic remap of "Driver" path to 32-bit equivalent
      if ($name -eq 'Driver' -and $value -is [string]) {
        $value = Convert-To32BitDriverPath $value
      }

      New-ItemProperty -Path $dstKey -Name $name -Value $value -PropertyType String -Force | Out-Null
    }

    # Update the 32-bit ODBC Data Sources list entry for this DSN to a valid 32-bit driver name if available
    $drvName32 = $drvName64
    $has32 = Test-32BitDriverExists $drvName64
    if (-not $has32) {
      # Some environments use slightly different friendly names between 64/32 bit. Try a few common tweaks.
      $candidateNames = @(
        $drvName64,
        ($drvName64 -replace '\s*\(x64\)\s*$', ''),
        ($drvName64 -replace '\s*64\-bit\s*$', ''),
        ($drvName64 -replace '\s*64 bit\s*$', ''),
        ($drvName64 -replace '\s*64\s*$', '')
      ) | Select-Object -Unique

      $drvName32 = $candidateNames | Where-Object { Test-32BitDriverExists $_ } | Select-Object -First 1
      if (-not $drvName32) {
        # Leave as-is but warn—ODBC admin may show DSN with a driver that's not installed in 32-bit.
        Write-Warning "32-bit driver not found for DSN '$dsnName' (frie ndly name '$drvName64'). DSN created but may not work until the 32-bit driver is installed."
        $drvName32 = $drvName64
      }
    }

    # Write the DSN entry under 32-bit ODBC Data Sources
    New-ItemProperty -Path $reg32_DSNList -Name $dsnName -Value $drvName32 -PropertyType String -Force | Out-Null

    $results += [PSCustomObject]@{
      DSN    = $dsnName
      Status = if ($exists32) { 'Overwritten' } else { 'Created' }
      Note   = if ($drvName32 -ne $drvName64) { "Driver mapped '$drvName64' -> '$drvName32'" } else { '' }
    }
  }
}

# Output a concise summary table
$results | Sort-Object DSN | Format-Table -AutoSize