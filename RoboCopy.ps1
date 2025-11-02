<# 
  RoboCopy.ps1
  Purpose: Menu-driven workstation migration assistant using Robocopy via admin shares (\\HOST\c$)

  Author: Andy Sendouw
  
  Highlights:
    - Ping test both endpoints before continuing
    - Compare (dry-run) -> parsed JSON
    - Copy non-system data
    - Copy Users\Public
    - Copy selected user profile(s)
    - Replay copy from JSON (live CLI feedback)
    - Toggle ACL mode (Inherit/Preserve)
    - Toggle threading (Fast=/MT:64, Streaming=/MT:1 + /ETA)
    - Toggle Skip OneDrive (exclude OneDrive/SkyDrive folders & *.cloud/*.cloudf)
    - Fix Security (ICACLS inheritance reset; optional grant)
    - Live console output via cmd /c + /TEE, while logging via /LOG+
    - Logs: \\fileserver01\Migration\Logs\(OLD to NEW)\logfile
#>

$Host.UI.RawUI.WindowTitle = "Robocopy Migration Helper"
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# ---------------- Paths & logging ----------------
function Get-LogRoot {
    param([string]$OldHost,[string]$NewHost)
    "\\fileserver01\Migration\Logs\($OldHost to $NewHost)\logfile"
}
function Ensure-Dir([string]$Path) {
    if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}
function New-LogFile {
    $root = Get-LogRoot -OldHost $script:OldHost -NewHost $script:NewHost
    Ensure-Dir $root
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    Join-Path $root "robocopy_$ts.log"
}

# ---------------- Exclusions ----------------
$Global:SystemDirExcludes = @(
  'Windows','Program Files','Program Files (x86)','Program Files (Arm)',
  'ProgramData','PerfLogs','Recovery','$Recycle.Bin','System Volume Information','OEM','Windows.old'
)
$Global:SystemFileExcludes = @('pagefile.sys','hiberfil.sys','swapfile.sys','swapfile.sys.tmp')

# ---------------- OneDrive skip (toggle) ----------------
$Global:SkipOneDrive = $true   # default ON

function Get-OneDriveDirs {
    param([string]$RootDrive) # e.g. \\OLD\c$  or  \\NEW\c$
    $dirs = @()
    $usersPath = Join-Path $RootDrive 'Users'
    if (-not (Test-Path $usersPath)) { return @() }

    $profiles = Get-ChildItem $usersPath -Directory -ErrorAction SilentlyContinue
    foreach ($p in $profiles) {
        try {
            $cand = Get-ChildItem -Path $p.FullName -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like 'OneDrive*' -or $_.Name -like 'SkyDrive*' } |
                    Select-Object -ExpandProperty FullName
            if ($cand) { $dirs += $cand }
        } catch { }
    }
    return ($dirs | Select-Object -Unique)
}

function Get-OneDriveFilePatterns {
    # OneDrive placeholders (avoid sparse/stub pulls)
    return @('*.cloud','*.cloudf')
}

# ---------------- Hostnames & connectivity ----------------
function Test-AdminShare([string]$Computer) { Test-Path "\\$Computer\c$" }

function Test-Ping {
    param([string]$Computer)
    Write-Host "Pinging $Computer..." -ForegroundColor Cyan
    try {
        if (Test-Connection -ComputerName $Computer -Count 2 -Quiet -ErrorAction Stop) {
            Write-Host "$Computer is reachable ✅" -ForegroundColor Green
            return $true
        } else {
            Write-Host "$Computer did not respond ❌" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Ping failed: $_" -ForegroundColor Red
        return $false
    }
}

function Require-Share([string]$Computer,[string]$Label) {
  if (-not (Test-AdminShare $Computer)) { throw "$Label admin share \\$Computer\c$ not reachable." }
}

function Read-Hostnames {
  Write-Host "`n=== Enter Hostnames ===" -ForegroundColor Cyan
  $script:OldHost = Read-Host "Old Computer Hostname"
  $script:NewHost = Read-Host "New Computer Hostname"

  if ([string]::IsNullOrWhiteSpace($OldHost) -or [string]::IsNullOrWhiteSpace($NewHost)) {
    throw "Hostnames cannot be blank."
  }

  if (-not (Test-Ping $OldHost)) { throw "Old computer $OldHost is unreachable. Aborting." }
  if (-not (Test-Ping $NewHost)) { throw "New computer $NewHost is unreachable. Aborting." }

  Require-Share $OldHost "Old"
  Require-Share $NewHost "New"

  $script:OldC = "\\$OldHost\c$"
  $script:NewC = "\\$NewHost\c$"
  Write-Host "OK: $OldC -> $NewC" -ForegroundColor Green
}

# ---------------- Modes ----------------
$Global:AclMode     = 'Inherit'   # 'Inherit' | 'Preserve'
$Global:ThreadMode  = 'Fast'      # 'Fast' (/MT:64) | 'Streaming' (/MT:1 + /ETA)

function Get-RoboCopyFlagsByAclMode {
  switch ($Global:AclMode) {
    'Preserve' { @('/COPY:DATS','/DCOPY:DAT','/SEC') }
    default    { @('/COPY:DAT','/DCOPY:DAT') }
  }
}
function Get-RoboCopyThreadFlags {
  switch ($Global:ThreadMode) {
    'Streaming' { @('/MT:1','/ETA') }
    default     { @('/MT:64') }  # DEFAULT NOW 64 THREADS
  }
}
function Get-ThreadModeDisplay {
  switch ($Global:ThreadMode) {
    'Streaming' { 'Streaming (/MT:1 + /ETA)' }
    default     { 'Fast (/MT:64)' }
  }
}
function Get-AclModeDisplay {
  $acl = Get-RoboCopyFlagsByAclMode -ErrorAction SilentlyContinue
  if (-not $acl) { return $Global:AclMode }
  return "$Global:AclMode  [" + ($acl -join ' ') + "]"
}

# --------- Argument explainer (maps switches -> human text) ----------
function Explain-RoboArgs {
  param([string[]]$Args)

  $explanations = New-Object System.Collections.Generic.List[string]

  foreach ($a in $Args) {
    switch -Regex ($a) {
      '^/E$'         { $explanations.Add('/E            – Include subdirs (even empty)'); continue }
      '^/R:(\d+)$'   { $explanations.Add("/R:$($Matches[1])       – Retry count per file"); continue }
      '^/W:(\d+)$'   { $explanations.Add("/W:$($Matches[1])       – Wait seconds between retries"); continue }
      '^/Z$'         { $explanations.Add('/Z            – Restartable mode'); continue }
      '^/V$'         { $explanations.Add('/V            – Verbose output'); continue }
      '^/TEE$'       { $explanations.Add('/TEE          – Mirror output to console + log'); continue }
      '^/XJ$'        { $explanations.Add('/XJ           – Exclude junctions/reparse points'); continue }
      '^/XN$'        { $explanations.Add('/XN           – Skip newer files on destination'); continue }
      '^/XO$'        { $explanations.Add('/XO           – Skip older files on destination'); continue }
      '^/XC$'        { $explanations.Add('/XC           – Skip changed files (no overwrite)'); continue }
      '^/COPY:DATS$' { $explanations.Add('/COPY:DATS    – Copy Data, Attributes, Timestamps, Security'); continue }
      '^/COPY:DAT$'  { $explanations.Add('/COPY:DAT     – Copy Data, Attributes, Timestamps'); continue }
      '^/DCOPY:DAT$' { $explanations.Add('/DCOPY:DAT    – Copy dir Data, Attributes, Timestamps'); continue }
      '^/SEC$'       { $explanations.Add('/SEC          – Copy NTFS ACLs (security)'); continue }
      '^/MT:(\d+)$'  { 
                        $t = [int]$Matches[1]
                        $label = if ($t -eq 1) { 'Streaming (single thread)' } else { "Fast multi-threading ($t threads)" }
                        $explanations.Add(("/MT:{0}      – {1}" -f $t,$label)); 
                        continue 
                     }
      '^/ETA$'       { $explanations.Add('/ETA          – Show estimated time of arrival'); continue }
      '^/FFT$'       { $explanations.Add('/FFT          – FAT file times (2-sec granularity)'); continue }
      '^/L$'         { $explanations.Add('/L            – List only (dry-run)'); continue }
      '^/NJH$'       { $explanations.Add('/NJH          – No job header'); continue }
      '^/NJS$'       { $explanations.Add('/NJS          – No job summary'); continue }
      '^/NS$'        { $explanations.Add('/NS           – No size'); continue }
      '^/NC$'        { $explanations.Add('/NC           – No class'); continue }
      '^/NP$'        { $explanations.Add('/NP           – No progress'); continue }
      '^/LOG\+:'     { $explanations.Add('/LOG+         – Append log output to file'); continue }
      '^/XD$'        { $explanations.Add('/XD           – Exclude the following directories (listed below)'); continue }
      '^/XF$'        { $explanations.Add('/XF           – Exclude the following files/patterns (listed below)'); continue }
      default        { }
    }
  }

  # Gather items following /XD and /XF to show them explicitly
  $xdItems = @()
  $xfItems = @()
  for ($i=0; $i -lt $Args.Count; $i++) {
    if ($Args[$i] -eq '/XD') {
      for ($j=$i+1; $j -lt $Args.Count -and ($Args[$j] -notmatch '^/'); $j++) { $xdItems += $Args[$j] }
    }
    if ($Args[$i] -eq '/XF') {
      for ($j=$i+1; $j -lt $Args.Count -and ($Args[$j] -notmatch '^/'); $j++) { $xfItems += $Args[$j] }
    }
  }
  if ($xdItems.Count -gt 0) {
    $explanations.Add("  XD items: " + (($xdItems | Select-Object -Unique) -join '; '))
  }
  if ($xfItems.Count -gt 0) {
    $explanations.Add("  XF items: " + (($xfItems | Select-Object -Unique) -join '; '))
  }

  return $explanations
}

# ---------------- Robocopy wrapper (LIVE console output) ----------------
function Invoke-Robo {
  param(
    [Parameter(Mandatory)] [string]$Source,
    [Parameter(Mandatory)] [string]$Dest,
    [string[]]$ExtraArgs = @(),
    [string]$LogFile
  )

  $base = @(
    $Source,$Dest,
    '/E','/R:1','/W:2',
    '/Z',
    '/V',
    '/TEE',             # mirror to console + log (live)
    '/XJ'               # skip junctions (e.g., profile reparse points)
  )
  $noOverwrite = @('/XN','/XO','/XC')
  $acl    = Get-RoboCopyFlagsByAclMode
  $thread = Get-RoboCopyThreadFlags

  $args = @(); $args += $base; $args += $thread; $args += $acl; $args += $noOverwrite
  if ($ExtraArgs) { $args += $ExtraArgs }
  if ($LogFile)   { $args += @("/LOG+:$LogFile") }

  # Quote and run via cmd for streaming; avoid UNC cwd issues
  $quoted  = $args | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }
  $argLine = $quoted -join ' '
  $safeWorkDir = "C:\Windows\System32"

  # Print exact command and a readable explainer
  Write-Host "`n>>> robocopy $argLine" -ForegroundColor Magenta
  $explain = Explain-RoboArgs -Args $args
  if ($explain.Count -gt 0) {
    Write-Host "    Switch explanations:" -ForegroundColor DarkCyan
    foreach ($line in $explain) { Write-Host "    $line" -ForegroundColor DarkGray }
  }

  Push-Location $safeWorkDir
  & cmd.exe /c "robocopy $argLine"
  $exit = $LASTEXITCODE
  Pop-Location

  Write-Host "Robocopy exit code: $exit" -ForegroundColor Yellow
  return $exit
}

# ---------------- Operations ----------------
function Compare-Tree {
    param([string]$Source,[string]$Dest,[string[]]$ExclusionsDirs,[string[]]$ExclusionsFiles)

    $log = New-LogFile
    $txt = [System.IO.Path]::ChangeExtension($log, ".txt")
    $json = [System.IO.Path]::ChangeExtension($log, ".json")

    $extra = @('/L','/FFT')
    if ($ExclusionsDirs)  { $extra += @('/XD') + $ExclusionsDirs }
    if ($ExclusionsFiles) { $extra += @('/XF') + $ExclusionsFiles }

    # Skip OneDrive (dirs + placeholders)
    if ($Global:SkipOneDrive) {
        $odDirs = Get-OneDriveDirs -RootDrive $Source
        if ($odDirs.Count -gt 0) { $extra += @('/XD') + $odDirs }
        $odFiles = Get-OneDriveFilePatterns
        if ($odFiles.Count -gt 0) { $extra += @('/XF') + $odFiles }
    }

    Write-Host "Running dry-run compare... (output will be parsed)" -ForegroundColor Cyan
    $listArgs = @($Source,$Dest) + $extra + @('/E','/R:0','/W:0','/V','/NJH','/NJS','/NS','/NC','/NP','/XJ')
    $qList = $listArgs | ForEach-Object { if ($_ -match '\s') { '"' + $_ + '"' } else { $_ } }
    & cmd.exe /c "robocopy $($qList -join ' ')" | Tee-Object -FilePath $txt

    Write-Host "Parsing dry-run results into JSON..." -ForegroundColor Yellow
    $pattern = '^(New File|Older|Extra File|Extra Dir|Newer)\s+([\d\.]+\w?)?\s+(.+)$'
    $results = @()
    Get-Content $txt | ForEach-Object {
        if ($_ -match $pattern) {
            # If skip OneDrive, don't include its entries in JSON either
            if ($Global:SkipOneDrive -and ($matches[3] -match '\\OneDrive( - |\\)|\\SkyDrive(\\|$)')) { return }
            $results += [PSCustomObject]@{
                Type = $matches[1]
                Size = $matches[2]
                Path = $matches[3]
            }
        }
    }

    $data = [PSCustomObject]@{
        Source = $Source
        Destination = $Dest
        Timestamp = (Get-Date)
        Differences = $results
    }

    $data | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 $json
    Write-Host "Dry-run comparison complete:`n - Raw: $txt`n - Parsed JSON: $json`n - Log: $log" -ForegroundColor Green
}

function Copy-NonSystemFromC {
  $log = New-LogFile
  $extra = @('/FFT','/XD') + $Global:SystemDirExcludes
  if ($Global:SystemFileExcludes.Count) { $extra += @('/XF') + $Global:SystemFileExcludes }

  if ($Global:SkipOneDrive) {
    $odDirs = Get-OneDriveDirs -RootDrive $script:OldC
    if ($odDirs.Count -gt 0) { $extra += @('/XD') + $odDirs }
    $odFiles = Get-OneDriveFilePatterns
    if ($odFiles.Count -gt 0) { $extra += @('/XF') + $odFiles }
  }

  Invoke-Robo -Source $script:OldC -Dest $script:NewC -ExtraArgs $extra -LogFile $log | Out-Null
  Write-Host "Copy complete. Log: $log" -ForegroundColor Green
}

function Copy-PublicProfile {
  $src = Join-Path $script:OldC 'Users\Public'
  $dst = Join-Path $script:NewC 'Users\Public'
  $log = New-LogFile
  Invoke-Robo -Source $src -Dest $dst -ExtraArgs @('/FFT') -LogFile $log | Out-Null
  Write-Host "Public profile copied. Log: $log" -ForegroundColor Green
}

function Select-UserProfiles {
  $usersPath = Join-Path $script:OldC 'Users'
  if (-not (Test-Path $usersPath)) { throw "Cannot access $usersPath" }
  $skip = @('Public','Default','Default User','All Users','defaultuser0','WDAGUtilityAccount')
  $dirs = Get-ChildItem $usersPath -Directory -ErrorAction SilentlyContinue | Where-Object { $skip -notcontains $_.Name }
  if (-not $dirs) { Write-Host "No user profiles found." -ForegroundColor Yellow; return @() }

  Write-Host "`nSelect user profile(s) to copy from \\$OldHost\c$\Users (comma-separated numbers)."
  for ($i=0; $i -lt $dirs.Count; $i++) { "{0,2}. {1}" -f ($i+1), $dirs[$i].Name }
  $raw = Read-Host "Enter numbers (e.g. 1,3,5) or 'A' for All"
  if ($raw -match '^[Aa]$') { return $dirs.Name }
  $idx = $raw -split '\s*,\s*' | ForEach-Object { [int]$_ - 1 } | Where-Object { $_ -ge 0 -and $_ -lt $dirs.Count }
  return $dirs[$idx].Name
}

function Copy-UserProfiles {
  $chosen = Select-UserProfiles
  if (-not $chosen -or $chosen.Count -eq 0) { Write-Host "No profiles selected; only Public will be copied." -ForegroundColor Yellow; return }
  foreach ($name in $chosen) {
    $src = Join-Path $script:OldC ("Users\$name")
    $dst = Join-Path $script:NewC ("Users\$name")
    $log = New-LogFile

    $extra = @('/FFT')
    if ($Global:SkipOneDrive) {
        $odDir = Get-OneDriveDirs -RootDrive $script:OldC | Where-Object { $_ -match "\\Users\\$([regex]::Escape($name))\\OneDrive" -or $_ -match "\\Users\\$([regex]::Escape($name))\\SkyDrive" }
        if ($odDir) { $extra += @('/XD') + $odDir }
        $odFiles = Get-OneDriveFilePatterns
        if ($odFiles.Count -gt 0) { $extra += @('/XF') + $odFiles }
    }

    Invoke-Robo -Source $src -Dest $dst -ExtraArgs $extra -LogFile $log | Out-Null
    Write-Host "Copied profile '$name'. Log: $log" -ForegroundColor Green
  }
}

# ---------------- Copy from JSON (with live CLI output) ----------------
function Copy-FromJson {
    $path = Read-Host "Enter path to JSON file from Compare (e.g. C:\Temp\robocopy_*.json)"
    if (-not (Test-Path $path)) { Write-Host "File not found: $path" -ForegroundColor Red; return }

    Write-Host "Loading JSON data..." -ForegroundColor Cyan
    $data = Get-Content $path -Raw | ConvertFrom-Json
    $log = New-LogFile

    # OneDrive folder list from Source to help skip in JSON replay
    $odRootDirs = if ($Global:SkipOneDrive) { Get-OneDriveDirs -RootDrive $data.Source } else { @() }

    foreach ($item in $data.Differences) {
        if ($item.Type -in @('New File','Newer','Older')) {

            # Skip OneDrive items when toggled on
            if ($Global:SkipOneDrive) {
                $p = $item.Path
                $isOD = $false
                if ($p -match '\\OneDrive( - |\\)|\\SkyDrive(\\|$)') { $isOD = $true }
                if (-not $isOD -and $odRootDirs.Count -gt 0) {
                    foreach ($od in $odRootDirs) {
                        if (("$($data.Source)\$p") -like "$od*") { $isOD = $true; break }
                    }
                }
                if ($isOD) { continue }
                if ($p -like '*.cloud' -or $p -like '*.cloudf') { continue }
            }

            $srcFile = Join-Path $data.Source $item.Path
            $dstFile = Join-Path $data.Destination $item.Path
            $dstDir  = Split-Path $dstFile -Parent
            if (-not (Test-Path $dstDir)) { New-Item -ItemType Directory -Path $dstDir -Force | Out-Null }

            try {
                Copy-Item -Path $srcFile -Destination $dstFile -Force -ErrorAction Stop
                Write-Host "✅ Copied: $($item.Path)" -ForegroundColor Green
                [System.Console]::Out.Flush()
                Add-Content $log "Copied: $srcFile -> $dstFile"
            } catch {
                Write-Host "❌ Failed: $($item.Path)" -ForegroundColor Red
                [System.Console]::Out.Flush()
                Add-Content $log "Failed: $srcFile -> $dstFile ($_)"
            }
        }
    }
    Write-Host "Selective copy complete. Log: $log" -ForegroundColor Green
}

# ---------------- Security fix-up ----------------
function Fix-Security {
  Write-Host "`n=== Fix Security (ICACLS) ===" -ForegroundColor Cyan
  $path = Read-Host "Enter destination path to fix (e.g. \\$NewHost\c$\Users or specific folder)"
  if ([string]::IsNullOrWhiteSpace($path)) { Write-Host "Cancelled." ; return }
  $grant = Read-Host "Grant Full Control to a user (e.g. NEWPC\username or DOMAIN\username) — or leave blank to skip"
  try {
    & icacls "$path" /inheritance:e /t | Out-Null
    if ($grant) {
      & icacls "$path" /grant:r "$grant`:(OI`)(CI`)F" /t | Out-Null
    }
    Write-Host "Security fixed for $path." -ForegroundColor Green
  } catch {
    Write-Host "ICACLS error: $_" -ForegroundColor Red
  }
}

# ---------------- Menu ----------------
function Show-Menu {
  $threadFlags = (Get-RoboCopyThreadFlags) -join ' '
  $aclFlags    = (Get-RoboCopyFlagsByAclMode) -join ' '
@"
==========================================
 Robocopy Migration Helper
 Old: $script:OldHost
 New: $script:NewHost

 Modes/Settings:
   ACL Mode:      $(Get-AclModeDisplay)
   Thread Mode:   $(Get-ThreadModeDisplay)  [$threadFlags]
   Skip OneDrive: $($Global:SkipOneDrive)

 Effective Defaults:
   Thread Flags:  $threadFlags
   ACL Flags:     $aclFlags

 Log Root:
   $(Get-LogRoot -OldHost $script:OldHost -NewHost $script:NewHost)
==========================================
1) Compare OLD C:\ to NEW C:\   (dry-run, excludes system + creates JSON)
2) Copy non-system data from OLD C:\ → NEW C:\
3) Copy Users\Public
4) Copy selected user profile(s)
5) Copy from JSON (selective replay with live output)
6) Change hostnames
7) Toggle ACL Mode (Inherit ↔ Preserve)
8) Toggle Thread Mode (Fast ↔ Streaming)
9) Fix Security (reset inheritance; optional grant)
10) Toggle Skip OneDrive (On/Off)
Q) Quit
"@
}

if (-not $script:OldHost -or -not $script:NewHost) { Read-Hostnames }

while ($true) {
  Show-Menu
  $choice = Read-Host "Choose an option"
  switch ($choice.ToUpper()) {
    '1'  { Compare-Tree -Source $script:OldC -Dest $script:NewC -ExclusionsDirs $Global:SystemDirExcludes -ExclusionsFiles $Global:SystemFileExcludes; Read-Host "Press ENTER to continue" | Out-Null }
    '2'  { Copy-NonSystemFromC;  Read-Host "Press ENTER to continue" | Out-Null }
    '3'  { Copy-PublicProfile;   Read-Host "Press ENTER to continue" | Out-Null }
    '4'  { Copy-UserProfiles;    Read-Host "Press ENTER to continue" | Out-Null }
    '5'  { Copy-FromJson;        Read-Host "Press ENTER to continue" | Out-Null }
    '6'  { Read-Hostnames }
    '7'  { $Global:AclMode    = if ($Global:AclMode -eq 'Inherit') { 'Preserve' } else { 'Inherit' };   Write-Host "ACL Mode is now: $(Get-AclModeDisplay)" -ForegroundColor Yellow }
    '8'  { $Global:ThreadMode = if ($Global:ThreadMode -eq 'Fast') { 'Streaming' } else { 'Fast' };     Write-Host "Thread Mode is now: $(Get-ThreadModeDisplay)" -ForegroundColor Yellow }
    '9'  { Fix-Security }
    '10' { $Global:SkipOneDrive = -not $Global:SkipOneDrive; Write-Host "Skip OneDrive is now: $Global:SkipOneDrive" -ForegroundColor Yellow }
    'Q'  { break }
    Default { Write-Host "Invalid choice." -ForegroundColor Yellow }
  }
}