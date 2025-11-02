<# 
  USMT-Menu.ps1
  Author: ANDY SENDOUW
  Menu-driven USMT capture/restore
#>

$USMTRootDefault  = '\\fileserver01\Migration\USMT\amd64'
$StoreRootDefault = '\\fileserver01\Migration\Stores'
$USMTExcludeBuiltins = @('/ue:*\Administrator','/ue:*\Default','/ue:*\Public','/ue:*\Guest')
$USMTExcludeAllThenBuiltins = @('/ue:*\*') + $USMTExcludeBuiltins

function PresentOrMissing([string]$Path) { 
  if (Test-Path $Path) { return (Resolve-Path -LiteralPath $Path).Path } 
  else { return "(missing)" } 
}

function PresentText([string]$Path) { 
  if (Test-Path $Path) { "present" } else { "missing" } 
}

function Write-Section($Text) {
  Write-Host ('=' * 70) -ForegroundColor DarkGray
  Write-Host $Text -ForegroundColor Cyan
  Write-Host ('=' * 70) -ForegroundColor DarkGray
}

function Test-PathOrThrow([string]$Path, [string]$What) {
  if (-not (Test-Path $Path)) { throw "$What not found: $Path" }
  return $true
}

function Ensure-Directory([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
  return $Path
}

function Get-USMTPath([string]$USMTRoot) {
  if ([string]::IsNullOrWhiteSpace($USMTRoot)) { $USMTRoot = $USMTRootDefault }
  if (Test-Path (Join-Path $USMTRoot 'scanstate.exe')) { return $USMTRoot }
  $arch = if ([Environment]::Is64BitOperatingSystem) { 'amd64' } else { 'x86' }
  $guess = Join-Path $USMTRoot $arch
  if (Test-Path (Join-Path $guess 'scanstate.exe')) { return $guess }
  throw "USMT binaries not found under: $USMTRoot"
}

function Get-StorePath([string]$StoreRoot, [string]$ComputerName) {
  if ([string]::IsNullOrWhiteSpace($StoreRoot)) { $StoreRoot = $StoreRootDefault }
  if ([string]::IsNullOrWhiteSpace($ComputerName)) { $ComputerName = $env:COMPUTERNAME }
  $p = Join-Path $StoreRoot $ComputerName
  Ensure-Directory $p | Out-Null
  return $p
}

function Get-LocalProfileTable {
  $profiles = Get-CimInstance -ClassName Win32_UserProfile |
    Where-Object { $_.LocalPath -and (Test-Path $_.LocalPath) -and -not $_.Special } |
    ForEach-Object {
      $sid  = $_.SID
      $path = $_.LocalPath
      try {
        $ntacct = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value
      } catch { $ntacct = $null }
      [pscustomobject]@{ SID = $sid; Account = $ntacct; Path = $path; LastUse = $_.LastUseTime }
    } |
    Where-Object {
      -not $_.Account -or (
        ($_.Account -notmatch '^(NT AUTHORITY|BUILTIN)\\') -and
        ($_.Account -notmatch '\\(Administrator|DefaultAccount|Guest|defaultuser0|WDAGUtilityAccount)$')
      )
    } |
    Sort-Object -Property LastUse -Descending

  $i = 1
  foreach ($p in $profiles) { Add-Member -InputObject $p -NotePropertyName Index -NotePropertyValue $i -Force; $i++ }
  return $profiles
}

function Show-ProfilePicker($Profiles) {
  $ProfilesArray = @($Profiles)
  if ($ProfilesArray.Count -eq 0) { Write-Warning "No eligible local profiles found."; return @() }

  Write-Host ""; Write-Host "Select profile(s) to CAPTURE (comma-separated indices), or choose:" -ForegroundColor Yellow
  Write-Host "  0) ALL user profiles (recommended for full machines)"; Write-Host ""
  
  $ProfilesArray | ForEach-Object {
    $acct = if ($_.Account) { $_.Account } else { "(SID-only) $($_.SID)" }
    $lastUse = if ($_.LastUse) { $_.LastUse.ToString('yyyy-MM-dd') } else { "never" }
    Write-Host ("{0,3}) {1,-40}  Last: {2}" -f $_.Index, $acct, $lastUse)
  }
  Write-Host ""

  $sel = (Read-Host "Your selection (e.g. 0 or 1,3,7)").Trim()
  if ([string]::IsNullOrWhiteSpace($sel)) { Write-Host "No selection made." -ForegroundColor Yellow; return @() }
  if ($sel -match '^(0|A|ALL)$') { Write-Host "Selected: ALL profiles" -ForegroundColor Green; return 'SELECTALL' }

  $validIndices = @()
  $tokens = $sel -split '[,\s]+'
  foreach ($token in $tokens) {
    $token = $token.Trim()
    if ([string]::IsNullOrEmpty($token)) { continue }
    if ($token -match '^\d+$') {
      $idx = [int]$token
      if ($idx -ge 1 -and $idx -le $ProfilesArray.Count) { $validIndices += $idx }
      else { Write-Host "Warning: Index $token is out of range (valid: 1-$($ProfilesArray.Count))" -ForegroundColor Yellow }
    } else { Write-Host "Warning: '$token' is not a valid number" -ForegroundColor Yellow }
  }
  
  if ($validIndices.Count -eq 0) { Write-Warning "No valid indices entered."; return @() }
  $validIndices = $validIndices | Sort-Object -Unique
  $picked = foreach ($i in $validIndices) { $ProfilesArray[$i - 1] }
  
  Write-Host "Selected $($picked.Count) profile(s):" -ForegroundColor Green
  $picked | ForEach-Object { 
    $acct = if ($_.Account) { $_.Account } else { $_.SID }
    Write-Host "  - $acct" -ForegroundColor Gray
  }
  return $picked
}

function Build-UIArgs($PickedProfiles) {
  if (-not $PickedProfiles) { return @() }
  $items = @($PickedProfiles)
  if ($items.Count -eq 1 -and $items[0] -is [string] -and $items[0] -eq 'SELECTALL') { return @('/all') }
  if ($items -contains 'SELECTALL') { return @('/all') }

  $selected = $items | ForEach-Object {
    if ($_ -is [string] -and $_ -eq 'SELECTALL') { return }
    $acct = $_.Account; $sid = $_.SID
    if ($acct -and ($acct -match '^[^\\]+\\[^\\]+$') -and ($acct -notmatch '^(NT AUTHORITY|BUILTIN)\\') -and 
        ($acct -notmatch '\\(Administrator|DefaultAccount|Guest|defaultuser0|WDAGUtilityAccount)$')) { "/ui:$acct" }
    elseif ($sid -and ($sid -match '^S-1-5-21-\d+-\d+-\d+-\d+$')) { "/ui:$sid" }
  } | Where-Object { $_ } | Sort-Object -Unique
  return $selected
}

function Show-StringPicker([string[]]$Items, [string]$PromptTitle) {
  if (-not $Items -or $Items.Count -eq 0) { Write-Warning "No items available to select."; return @() }
  Write-Host ""; Write-Host $PromptTitle -ForegroundColor Yellow
  Write-Host "  0) ALL (restore everything in this store)"; Write-Host ""
  for ($i=0; $i -lt $Items.Count; $i++) { Write-Host ("{0,3}) {1}" -f ($i+1), $Items[$i]) }
  Write-Host ""
  
  $sel = (Read-Host "Your selection (e.g. 0 or 1,3,7)").Trim()
  if ([string]::IsNullOrWhiteSpace($sel)) { Write-Host "No selection made." -ForegroundColor Yellow; return @() }
  if ($sel -match '^(0|A|ALL)$') { Write-Host "Selected: ALL" -ForegroundColor Green; return 'SELECTALL' }
  
  $validIndices = @()
  $tokens = $sel -split '[,\s]+'
  foreach ($token in $tokens) {
    $token = $token.Trim()
    if ([string]::IsNullOrEmpty($token)) { continue }
    if ($token -match '^\d+$') {
      $idx = [int]$token
      if ($idx -ge 1 -and $idx -le $Items.Count) { $validIndices += $idx }
      else { Write-Host "Warning: Index $token is out of range (1-$($Items.Count))" -ForegroundColor Yellow }
    } else { Write-Host "Warning: '$token' is not a valid number" -ForegroundColor Yellow }
  }
  
  if ($validIndices.Count -eq 0) { Write-Warning "No valid selections."; return @() }
  $validIndices = $validIndices | Sort-Object -Unique
  $picked = foreach ($i in $validIndices) { $Items[$i-1] }
  Write-Host "Selected $($picked.Count) item(s)" -ForegroundColor Green
  return ,$picked
}

function Build-ScanArgs([string]$USMTRoot, [string]$StorePath, [string[]]$UIArgs, [string]$Key, [int]$UelDays = 0) {
  $MigDocs = Join-Path $USMTRoot 'migdocs.xml'; $MigApp = Join-Path $USMTRoot 'migapp.xml'
  $LogFile = Join-Path $StorePath 'scan.log'; $ProgFile = Join-Path $StorePath 'scan.progress.log'
  $args = @("`"$StorePath`"",'/o','/c','/v:13','/vsc',"/l:`"$LogFile`"","/progress:`"$ProgFile`"",
            "/listfiles:`"$(Join-Path $StorePath 'scan.files.txt')`"","/i:`"$MigDocs`"","/i:`"$MigApp`"")
  $usingAll = $UIArgs -and ($UIArgs -contains '/all')
  if ($usingAll) { $args += $USMTExcludeBuiltins + '/all' }
  else { $args += $USMTExcludeAllThenBuiltins; if ($UIArgs -and $UIArgs.Count -gt 0) { $args += $UIArgs } }
  if ($Key) { $args += "/key:`"$Key`"" }
  if ($UelDays -gt 0) { $args += "/uel:$UelDays" }
  $modeFile = Join-Path $StorePath 'scan.identity.mode.txt'
  try { $mode = if ($usingAll) { 'ALL' } else { 'STRICT' }; Set-Content -LiteralPath $modeFile -Value $mode -Encoding ASCII } catch {}
  return $args
}

function Build-LoadArgs([string]$USMTRoot, [string]$StorePath, [hashtable]$Mappings, [string[]]$UIArgs, [string]$Key) {
  $MigDocs = Join-Path $USMTRoot 'migdocs.xml'; $MigApp = Join-Path $USMTRoot 'migapp.xml'
  $LogFile = Join-Path $StorePath 'load.log'; $ProgFile = Join-Path $StorePath 'load.progress.log'
  $args = @("`"$StorePath`"",'/c','/v:13','/vsc',"/l:`"$LogFile`"","/progress:`"$ProgFile`"","/i:`"$MigDocs`"","/i:`"$MigApp`"")
  $usingAll = $UIArgs -and ($UIArgs -contains '/all')
  if ($usingAll) { $args += '/all' }
  elseif ($UIArgs -and $UIArgs.Count -gt 0) { $args += @('/ue:*\*') + $UIArgs }
  if ($Mappings -and $Mappings.Count -gt 0) {
    foreach ($k in $Mappings.Keys) { $v = $Mappings[$k]; if ($k -and $v) { $args += "/mu:`"$k`"=`"$v`"" } }
  }
  if ($Key) { $args += "/key:`"$Key`"" }
  return $args
}

function Invoke-USMTWithProgress {
  param([Parameter(Mandatory)][string]$ExePath, [Parameter(Mandatory)][string[]]$Arguments,
        [Parameter(Mandatory)][string]$ProgressFile, [string]$Activity = "USMT", [string]$LogFile)
  
  Write-Host "`nStarting USMT operation..." -ForegroundColor Cyan
  Write-Host "Progress will be displayed below:" -ForegroundColor Gray; Write-Host ""
  if (Test-Path $ProgressFile) { Remove-Item $ProgressFile -Force -ErrorAction SilentlyContinue }
  
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $ExePath; $psi.Arguments = $Arguments -join ' '; $psi.UseShellExecute = $false
  $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.CreateNoWindow = $true
  $p = New-Object System.Diagnostics.Process; $p.StartInfo = $psi
  
  $stdoutHandler = { if (-not [string]::IsNullOrWhiteSpace($EventArgs.Data)) { Write-Host $EventArgs.Data -ForegroundColor Gray } }
  $stderrHandler = { if (-not [string]::IsNullOrWhiteSpace($EventArgs.Data)) { Write-Host $EventArgs.Data -ForegroundColor Yellow } }
  Register-ObjectEvent -InputObject $p -EventName OutputDataReceived -Action $stdoutHandler | Out-Null
  Register-ObjectEvent -InputObject $p -EventName ErrorDataReceived -Action $stderrHandler | Out-Null
  
  try {
    $p.Start() | Out-Null; $p.BeginOutputReadLine(); $p.BeginErrorReadLine()
    $lastPct = -1; $lastStatus = ""
    while (-not $p.HasExited) {
      Start-Sleep -Milliseconds 500
      if (Test-Path $ProgressFile) {
        try {
          $line = Get-Content -LiteralPath $ProgressFile -Tail 1 -ErrorAction SilentlyContinue
          if ($line -and $line -match '\((\d+)%\)') {
            $pct = [int]$matches[1]
            if ($pct -ne $lastPct) {
              $lastPct = $pct; Write-Progress -Activity $Activity -PercentComplete $pct -Status $line
              if ($pct % 10 -eq 0 -and $line -ne $lastStatus) { Write-Host "[$pct%] $line" -ForegroundColor Cyan; $lastStatus = $line }
            }
          }
        } catch {}
      }
    }
  } finally {
    Write-Progress -Activity $Activity -Completed; $p.WaitForExit()
    Get-EventSubscriber | Where-Object { $_.SourceObject -eq $p } | Unregister-Event
    Write-Host ""
  }
  return $p.ExitCode
}

function Action-Capture {
  try {
    Write-Section "CAPTURE (ScanState)"
    $USMTRoot = Get-USMTPath -USMTRoot $USMTRootDefault
    $StorePath = Get-StorePath -StoreRoot $StoreRootDefault -ComputerName $env:COMPUTERNAME
    $ScanExe = Join-Path $USMTRoot 'scanstate.exe'
    Test-PathOrThrow $ScanExe "ScanState"; Ensure-Directory $StorePath | Out-Null
    Write-Host "Computer: $env:COMPUTERNAME" -ForegroundColor Cyan
    Write-Host "Store: $StorePath" -ForegroundColor Cyan; Write-Host ""
    $profiles = Get-LocalProfileTable
    if (-not $profiles -or $profiles.Count -eq 0) { Write-Warning "No user profiles found on this computer."; return }
    Write-Host "Found $($profiles.Count) user profile(s)" -ForegroundColor Green
    $picked = Show-ProfilePicker -Profiles $profiles
    if (-not $picked -or $picked.Count -eq 0) { Write-Warning "Nothing selected. Returning to menu."; return }
    $uiArgs = Build-UIArgs -PickedProfiles $picked
    if (-not $uiArgs -or $uiArgs.Count -eq 0) { Write-Warning "No valid user identities after validation. Aborting."; return }
    Write-Host "`nUSMT identities to include:" -ForegroundColor Yellow
    $uiArgs | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
    $selJson = Join-Path $StorePath 'selected_identities.json'
    try {
      $manifest = [pscustomobject]@{ version = 2; generatedAt = (Get-Date).ToString('s'); sourceComputer = $env:COMPUTERNAME; identities = @($uiArgs) }
      $manifest | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $selJson -Encoding UTF8
      Write-Host "Saved selection manifest: $selJson" -ForegroundColor Green
    } catch { Write-Warning "Could not write JSON selection manifest: $_" }
    Write-Host ""; $uelInput = Read-Host "Limit to users active in the last N days? (Enter to skip)"; $uelDays = 0
    if (-not [string]::IsNullOrWhiteSpace($uelInput)) {
      if ([int]::TryParse($uelInput, [ref]$null)) { $uelDays = [int]$uelInput; Write-Host "Will capture only users active in last $uelDays days" -ForegroundColor Cyan }
      else { Write-Warning "Invalid number; ignoring /uel filter." }
    }
    $key = Read-Host "Optional encryption key (press Enter to skip)"
    if ($key) { Write-Host "Encryption enabled - remember this key for restore!" -ForegroundColor Yellow }
    $args = Build-ScanArgs -USMTRoot $USMTRoot -StorePath $StorePath -UIArgs $uiArgs -Key $key -UelDays $uelDays
    Write-Host "`n--- CAPTURE COMMAND ---" -ForegroundColor Cyan
    Write-Host "$ScanExe $($args -join ' ')" -ForegroundColor DarkGray; Write-Host ""
    $confirm = Read-Host "Proceed with capture? (Y/N)"
    if ($confirm -notmatch '^[Yy]') { Write-Host "Cancelled." -ForegroundColor Yellow; return }
    $scanLog = Join-Path $StorePath 'scan.log'
    $code = Invoke-USMTWithProgress -ExePath $ScanExe -Arguments $args -ProgressFile (Join-Path $StorePath 'scan.progress.log') -Activity "ScanState (capturing)" -LogFile $scanLog
    Write-Host "`n--- RESULTS ---" -ForegroundColor Cyan
    Write-Host "ScanState exit code: $code" -ForegroundColor $(if ($code -eq 0) { 'Green' } else { 'Yellow' })
    if ($code -eq 0) {
      Write-Host "Capture succeeded!" -ForegroundColor Green; Write-Host "Store location: $StorePath" -ForegroundColor Gray
      $filesList = Join-Path $StorePath 'scan.files.txt'
      if (Test-Path $filesList) { $count = (Get-Content -LiteralPath $filesList | Measure-Object -Line).Lines; Write-Host "Files captured: $count" -ForegroundColor Gray }
    } else { Write-Warning "Capture finished with non-zero exit code."; Write-Host "Check log: $scanLog" -ForegroundColor Yellow }
  } catch { Write-Error $_ }
}

function Action-Restore {
  try {
    Write-Section "RESTORE (LoadState)"
    $USMTRoot = Get-USMTPath -USMTRoot $USMTRootDefault; $LoadExe = Join-Path $USMTRoot 'loadstate.exe'
    Test-PathOrThrow $LoadExe "LoadState"
    Write-Host "Current computer: $env:COMPUTERNAME" -ForegroundColor Cyan; Write-Host ""
    $srcComp = Read-Host "Enter SOURCE computer name (default: $env:COMPUTERNAME)"
    if ([string]::IsNullOrWhiteSpace($srcComp)) { $srcComp = $env:COMPUTERNAME }
    $StorePath = Join-Path $StoreRootDefault $srcComp
    if (-not (Test-Path $StorePath)) {
      Write-Warning "Store path not found: $StorePath"; $create = Read-Host "Create directory anyway? (Y/N)"
      if ($create -match '^[Yy]') { Ensure-Directory $StorePath | Out-Null } else { return }
    }
    Write-Host "Store: $StorePath" -ForegroundColor Cyan
    $selJson = Join-Path $StorePath 'selected_identities.json'; $restoreUI = @(); $captured = @()
    if (Test-Path $selJson) {
      try {
        Write-Host "Found capture manifest" -ForegroundColor Green
        $mf = Get-Content -LiteralPath $selJson -Raw | ConvertFrom-Json
        if ($mf -and $mf.identities) {
          $captured = @($mf.identities); Write-Host "Captured from: $($mf.sourceComputer)" -ForegroundColor Gray
          Write-Host "Captured on: $($mf.generatedAt)" -ForegroundColor Gray
        }
      } catch { Write-Warning "Failed reading JSON manifest: $_" }
    }
    if ($captured) { $captured = $captured | Sort-Object -Unique }
    if ($captured -and $captured.Count -gt 0) {
      Write-Host "`nAvailable identities in store: $($captured.Count)" -ForegroundColor Cyan
      $picked = Show-StringPicker -Items $captured -PromptTitle "Select profile(s) to RESTORE from the captured set"
      if ($picked -eq 'SELECTALL') { $restoreUI = @('/all'); Write-Host "Will restore ALL captured profiles" -ForegroundColor Green }
      elseif ($picked -and $picked.Count -gt 0) { $restoreUI = @($picked); Write-Host "Will restore $($restoreUI.Count) specific profile(s)" -ForegroundColor Green }
      else { Write-Host "No selection made; defaulting to ALL from store." -ForegroundColor Yellow; $restoreUI = @('/all') }
    } else {
      Write-Warning "No JSON selection manifest found in store."; $ans = (Read-Host "Restore ALL users from store? (Y/N)").Trim()
      if ($ans -match '^[Yy]') { $restoreUI = @('/all') }
      else {
        Write-Host "Enter /ui filters manually (e.g. DOMAIN\User or SID); blank line to finish." -ForegroundColor Yellow
        while ($true) {
          $line = Read-Host "/ui"; if ([string]::IsNullOrWhiteSpace($line)) { break }
          if ($line -match '^/ui:.+') { $restoreUI += $line.Trim() }
          elseif ($line -notmatch '^/ui:') { $restoreUI += "/ui:$line" }
          else { Write-Warning "Invalid format" }
        }
      }
    }
    Write-Host "`nUser account mappings (optional):" -ForegroundColor Cyan
    Write-Host "Format: 'OLD\User=NEW\User' or 'OLD_SID=NEW\User'" -ForegroundColor Gray
    Write-Host "Enter one per line, blank line to finish" -ForegroundColor Gray
    $map = @{}
    while ($true) {
      $line = Read-Host "map"; if ([string]::IsNullOrWhiteSpace($line)) { break }
      if ($line -match '^[^=]+=[^=]+$') {
        $parts = $line -split '=',2; $map[$parts[0].Trim()] = $parts[1].Trim()
        Write-Host "  Added: $($parts[0]) -> $($parts[1])" -ForegroundColor Gray
      } else { Write-Warning "Invalid format. Example: DOMAIN\OldUser=DOMAIN\NewUser" }
    }
    $key = Read-Host "`nIf you used an encryption key during capture, enter it now (Enter to skip)"
    $args = Build-LoadArgs -USMTRoot $USMTRoot -StorePath $StorePath -Mappings $map -UIArgs $restoreUI -Key $key
    Write-Host "`n--- RESTORE COMMAND ---" -ForegroundColor Cyan
    Write-Host "$LoadExe $($args -join ' ')" -ForegroundColor DarkGray; Write-Host ""
    $confirm = Read-Host "Proceed with restore? (Y/N)"
    if ($confirm -notmatch '^[Yy]') { Write-Host "Cancelled." -ForegroundColor Yellow; return }
    $loadLog = Join-Path $StorePath 'load.log'
    $code = Invoke-USMTWithProgress -ExePath $LoadExe -Arguments $args -ProgressFile (Join-Path $StorePath 'load.progress.log') -Activity "LoadState (restoring)" -LogFile $loadLog
    Write-Host "`n--- RESULTS ---" -ForegroundColor Cyan
    Write-Host "LoadState exit code: $code" -ForegroundColor $(if ($code -eq 0) { 'Green' } else { 'Yellow' })
    if ($code -eq 0) { Write-Host "Restore succeeded!" -ForegroundColor Green; Write-Host "Data restored from: $StorePath" -ForegroundColor Gray }
    else { Write-Warning "Restore finished with non-zero exit code."; Write-Host "Check log: $loadLog" -ForegroundColor Yellow }
  } catch { Write-Error $_ }
}

function Action-ShowManifest {
  try {
    Write-Section "SHOW MANIFEST"
    $comp = Read-Host "Enter computer name (default: $env:COMPUTERNAME)"
    if ([string]::IsNullOrWhiteSpace($comp)) { $comp = $env:COMPUTERNAME }
    $storePath = Get-StorePath -StoreRoot $StoreRootDefault -ComputerName $comp
    $jsonPath = Join-Path $storePath 'selected_identities.json'; $modePath = Join-Path $storePath 'scan.identity.mode.txt'
    $filesPath = Join-Path $storePath 'scan.files.txt'; $scanLog = Join-Path $storePath 'scan.log'; $loadLog = Join-Path $storePath 'load.log'
    Write-Host ""; Write-Host "Store path  : $storePath" -ForegroundColor Cyan
    Write-Host "Manifest    : $(PresentOrMissing $jsonPath)" -ForegroundColor Gray
    Write-Host "Mode file   : $(PresentOrMissing $modePath)" -ForegroundColor Gray
    Write-Host "Files list  : $(PresentOrMissing $filesPath)" -ForegroundColor Gray; Write-Host ""
    if (Test-Path $jsonPath) {
      try {
        $mf = Get-Content -LiteralPath $jsonPath -Raw | ConvertFrom-Json
        Write-Host "--- MANIFEST DETAILS ---" -ForegroundColor Yellow
        Write-Host "Version         : $($mf.version)"; Write-Host "Generated       : $($mf.generatedAt)"; Write-Host "Source Computer : $($mf.sourceComputer)"
        if ($mf.identities -and $mf.identities.Count -gt 0) {
          Write-Host "Identities      : $($mf.identities.Count)"
          $mf.identities | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
        } else { Write-Host "Identities      : (none)" -ForegroundColor DarkGray }
      } catch { Write-Warning "Failed to parse manifest JSON: $_" }
    } else { Write-Warning "No selected_identities.json found." }
    Write-Host ""
    if (Test-Path $modePath) { $mode = (Get-Content -LiteralPath $modePath -TotalCount 1).Trim(); Write-Host "Identity mode: $mode" -ForegroundColor Cyan }
    if (Test-Path $filesPath) {
      $count = (Get-Content -LiteralPath $filesPath | Measure-Object -Line).Lines; Write-Host "Files listed : $count" -ForegroundColor Cyan
      if ($count -gt 0) { Write-Host "`nSample files (first 20):" -ForegroundColor Yellow; Get-Content -LiteralPath $filesPath -TotalCount 20 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray } }
    } else { Write-Host "No files list found." -ForegroundColor DarkGray }
    Write-Host "`n--- LOG FILES ---" -ForegroundColor Yellow
    Write-Host "scan.log : $(PresentText $scanLog)" -ForegroundColor Gray; Write-Host "load.log : $(PresentText $loadLog)" -ForegroundColor Gray
    if (Test-Path $storePath) {
      try {
        $size = (Get-ChildItem -Path $storePath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $sizeMB = [math]::Round($size / 1MB, 2); $sizeGB = [math]::Round($size / 1GB, 2)
        if ($sizeGB -gt 1) { Write-Host "Store size   : $sizeGB GB" -ForegroundColor Cyan }
        else { Write-Host "Store size   : $sizeMB MB" -ForegroundColor Cyan }
      } catch { Write-Host "Store size   : (unable to calculate)" -ForegroundColor DarkGray }
    }
  } catch { Write-Error $_ }
}

function Show-Menu {
  Clear-Host
  Write-Section "USMT Menu Tool"
  Write-Host "Computer     : $env:COMPUTERNAME" -ForegroundColor Gray
  Write-Host "USMT root    : $USMTRootDefault" -ForegroundColor Gray
  Write-Host "Store root   : $StoreRootDefault\<COMPUTERNAME>" -ForegroundColor Gray
  Write-Host ""
  Write-Host "1) Capture (ScanState)      - Save user profiles from this PC"
  Write-Host "2) Restore (LoadState)      - Restore user profiles to this PC"
  Write-Host "3) Show manifest / summary  - View capture details"
  Write-Host "4) Validate paths           - Check USMT installation"
  Write-Host "5) Quit"
  Write-Host ""
}

$__RunMenu = $true
while ($__RunMenu) {
  try {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
      '1' { Action-Capture; Write-Host ""; Read-Host "Press Enter to continue..." | Out-Null }
      '2' { Action-Restore; Write-Host ""; Read-Host "Press Enter to continue..." | Out-Null }
      '3' { Action-ShowManifest; Write-Host ""; Read-Host "Press Enter to continue..." | Out-Null }
      '4' {
        Write-Section "VALIDATION"
        try {
          Write-Host "Checking USMT installation..." -ForegroundColor Cyan
          $USMTRoot = Get-USMTPath -USMTRoot $USMTRootDefault
          $scanExe = Join-Path $USMTRoot 'scanstate.exe'
          $loadExe = Join-Path $USMTRoot 'loadstate.exe'
          Test-PathOrThrow $scanExe "ScanState"
          Test-PathOrThrow $loadExe "LoadState"
          Write-Host "USMT root: $USMTRoot" -ForegroundColor Green
          Write-Host "ScanState.exe found" -ForegroundColor Green
          Write-Host "LoadState.exe found" -ForegroundColor Green
          $migDocs = Join-Path $USMTRoot 'migdocs.xml'
          $migApp = Join-Path $USMTRoot 'migapp.xml'
          if (Test-Path $migDocs) { Write-Host "migdocs.xml found" -ForegroundColor Green }
          else { Write-Warning "migdocs.xml not found" }
          if (Test-Path $migApp) { Write-Host "migapp.xml found" -ForegroundColor Green }
          else { Write-Warning "migapp.xml not found" }
          Write-Host ""
          Write-Host "Checking store access..." -ForegroundColor Cyan
          $store = Get-StorePath -StoreRoot $StoreRootDefault -ComputerName $env:COMPUTERNAME
          if (Test-Path $store) {
            Write-Host "Store path exists: $store" -ForegroundColor Green
            $testFile = Join-Path $store '.test'
            try {
              "test" | Out-File -FilePath $testFile -Force
              Remove-Item $testFile -Force
              Write-Host "Store is writable" -ForegroundColor Green
            } catch { Write-Warning "Store may not be writable: $_" }
          } else {
            Write-Host "Store path does not exist (will be created on first capture)" -ForegroundColor Yellow
            Write-Host "Path: $store" -ForegroundColor Gray
          }
          Write-Host ""
          Write-Host "--- SYSTEM INFO ---" -ForegroundColor Cyan
          Write-Host "OS: $([Environment]::OSVersion.VersionString)"
          Write-Host "Architecture: $(if ([Environment]::Is64BitOperatingSystem) { 'x64' } else { 'x86' })"
          Write-Host "PowerShell: $($PSVersionTable.PSVersion)"
        } catch { Write-Error $_ }
        Write-Host ""
        Read-Host "Press Enter to continue..." | Out-Null
      }
      '5' { Write-Host "Exiting... Goodbye!" -ForegroundColor Cyan; $__RunMenu = $false }
      default { Write-Host "Invalid selection. Please choose 1-5." -ForegroundColor Red; Start-Sleep -Seconds 1 }
    }
  } catch {
    Write-Host "`nError: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press Enter to continue..." | Out-Null
  }
}
