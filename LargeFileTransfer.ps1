<# 
  Start-FolderTransfer.ps1 (Safe Input Edition)
  Robust Robocopy menu with input normalization and error handling
  Works over admin shares (C$, D$). No need for users to type slashes perfectly.
#>

$DefaultThreads = 16
$DefaultIpg     = 0
$LogsDir        = 'C:\Temp\robocopy_logs'
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null

function Normalize-SubPath($path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return "" }
    $path = $path -replace '/', '\'                   # Convert forward slashes
    $path = $path.Trim() -replace '^[\\]+', ''        # Remove leading slashes
    return $path.TrimEnd('\')                         # Remove trailing slashes
}

function Read-PathInputs {
    Write-Host "`n=== Enter Source & Destination ===" -ForegroundColor Cyan
    $srcHost = Read-Host "Source host (e.g., OLDHOST)"
    if ([string]::IsNullOrWhiteSpace($srcHost)) { throw "Source host cannot be blank." }

    $srcDrive = Read-Host "Source drive letter (default C)"
    if ([string]::IsNullOrWhiteSpace($srcDrive)) { $srcDrive = 'C' }
    $srcDrive = $srcDrive.TrimEnd(':')                # Strip colon if present

    $srcFolder = Read-Host "Source folder path on that drive (e.g., Users\Shared\BigFolder)"
    $srcFolder = Normalize-SubPath $srcFolder

    $dstHost = Read-Host "Destination host (e.g., NEWHOST)"
    if ([string]::IsNullOrWhiteSpace($dstHost)) { throw "Destination host cannot be blank." }

    $dstDrive = Read-Host "Destination drive letter (default C)"
    if ([string]::IsNullOrWhiteSpace($dstDrive)) { $dstDrive = 'C' }
    $dstDrive = $dstDrive.TrimEnd(':')

    $dstFolder = Read-Host "Destination folder path on that drive (e.g., Data\BigFolder)"
    $dstFolder = Normalize-SubPath $dstFolder

    $src = "\\$srcHost\$($srcDrive)$"
    $dst = "\\$dstHost\$($dstDrive)$"

    if ($srcFolder) { $src = Join-Path $src $srcFolder }
    if ($dstFolder) { $dst = Join-Path $dst $dstFolder }

    Write-Host "`nSource:      $src"
    Write-Host "Destination: $dst`n"

    # Test access
    try { if (-not (Test-Path $src)) { Write-Warning "⚠ Source not reachable: $src" } } catch {}
    try { if (-not (Test-Path $dst)) { Write-Warning "⚠ Destination not reachable: $dst" } catch {}

    return [PSCustomObject]@{ Src=$src; Dst=$dst }
}

function Read-PerfOptions {
    Write-Host "`n=== Performance Options ===" -ForegroundColor Cyan
    $mt = Read-Host "Threads (/MT:N) [default $DefaultThreads]"
    if (-not [int]::TryParse($mt, [ref]0)) { $mt = $DefaultThreads } else { $mt = [int]$mt }
    $ipg = Read-Host "Throttle Inter-Packet Gap ms (/IPG:N) [default $DefaultIpg]"
    if (-not [int]::TryParse($ipg, [ref]0)) { $ipg = $DefaultIpg } else { $ipg = [int]$ipg }
    $useBackupMode = Read-Host "Use backup mode (/B) to copy open files? [y/N]"
    $useBackupMode = $useBackupMode -match '^(y|yes)$'
    return [PSCustomObject]@{ MT=$mt; IPG=$ipg; Backup=$useBackupMode }
}

function New-LogFile([string]$prefix) {
    $runId = Get-Date -Format 'yyyyMMdd-HHmmss'
    return Join-Path $LogsDir "$prefix-$runId.log"
}

function Invoke-Robo {
    param(
        [Parameter(Mandatory)] [string] $Src,
        [Parameter(Mandatory)] [string] $Dst,
        [ValidateSet('DryRun','BulkCopy','Mirror')] [string] $Mode,
        [int] $MT = 16,
        [int] $IPG = 0,
        [switch] $BackupMode
    )

    $log = New-LogFile $Mode
    $args = @(
        "`"$Src`"", "`"$Dst`"",
        '/ETA','/Z','/J','/R:3','/W:5','/V','/TEE',
        "/LOG:`"$log`"","/MT:$MT"
    )
    if ($IPG -gt 0) { $args += "/IPG:$IPG" }
    if ($BackupMode) { $args += '/B' }

    switch ($Mode) {
        'DryRun'   { $args += '/E','/L' }
        'BulkCopy' { $args += '/E' }
        'Mirror'   {
            Write-Warning "⚠ MIRROR will delete extra files at destination."
            $confirm = Read-Host "Type MIRROR to confirm"
            if ($confirm -ne 'MIRROR') { Write-Host "Cancelled."; return }
            $args += '/MIR'
        }
    }

    Write-Host "`nRunning Robocopy ($Mode)...`n" -ForegroundColor Green
    & robocopy.exe @args
    $exit = $LASTEXITCODE
    Write-Host "`nExit Code: $exit (0–1 = success; >1 = partial failures)" -ForegroundColor Yellow
    Write-Host "Log saved: $log`n"
}

function Estimate-Size([string]$Src) {
    Write-Host "`nEstimating size of $Src ..." -ForegroundColor Cyan
    try {
        $items = Get-ChildItem -LiteralPath $Src -Recurse -File -ErrorAction Stop
        $size = ($items | Measure-Object Length -Sum).Sum
        Write-Host ("{0:N0} files, {1:N2} GB total" -f $items.Count, ($size/1GB))
    } catch {
        Write-Warning "Failed to enumerate: $($_.Exception.Message)"
    }
}

function Show-Menu {
@"
========================================
   Large Folder Transfer Utility
========================================
1) Dry Run (preview; no copy)
2) Bulk Copy (safe; no deletes)
3) Mirror (make dest match source)
4) Estimate source size
5) Exit
"@
}

# --- Main Loop ---
while ($true) {
    Show-Menu
    $choice = Read-Host "Choose an option (1-5)"
    switch ($choice) {
        '1' {
            $p = Read-PathInputs
            $o = Read-PerfOptions
            Invoke-Robo -Src $p.Src -Dst $p.Dst -Mode DryRun -MT $o.MT -IPG $o.IPG -BackupMode:$o.Backup
        }
        '2' {
            $p = Read-PathInputs
            $o = Read-PerfOptions
            Invoke-Robo -Src $p.Src -Dst $p.Dst -Mode BulkCopy -MT $o.MT -IPG $o.IPG -BackupMode:$o.Backup
        }
        '3' {
            $p = Read-PathInputs
            $o = Read-PerfOptions
            Invoke-Robo -Src $p.Src -Dst $p.Dst -Mode Mirror -MT $o.MT -IPG $o.IPG -BackupMode:$o.Backup
        }
        '4' {
            $p = Read-PathInputs
            Estimate-Size -Src $p.Src
        }
        '5' { break }
        default { Write-Host "Invalid choice." -ForegroundColor Red }
    }
}