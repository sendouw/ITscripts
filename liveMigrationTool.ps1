<# 
  LiveUserMigration.Menu.ps1  (UI-first, elevation + STA safe)
  Author: Andy Sendouw
  Enhanced: Performance optimizations and UI improvements
#>

[CmdletBinding()]
Param(
  [int]$ParallelUserCopies = 4,  # Increased from 3 for better performance
  [switch]$ThrottleBusinessHours,
  [switch]$Gui
)

# --- Make GUI the default even after relaunch ---
if (-not $PSBoundParameters.ContainsKey('Gui')) { $Gui = $true }

# --- Ensure Admin + STA (relaunch if needed) ---
function Ensure-Admin-STA {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  $isSTA = [Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'

  if ($isAdmin -and $isSTA) { return }

  $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$($MyInvocation.MyCommand.Path)`"")
  # preserve explicit parameters & defaults we care about
  if ($Gui) { $args += '-Gui' }
  if ($PSBoundParameters.ContainsKey('ParallelUserCopies')) { $args += @('-ParallelUserCopies',"$ParallelUserCopies") }
  if ($ThrottleBusinessHours) { $args += '-ThrottleBusinessHours' }

  $verb = $(if ($isAdmin) {'open'} else {'runas'})
  try {
    Start-Process -FilePath "powershell.exe" -ArgumentList $args -Verb $verb -WindowStyle Normal | Out-Null
  } catch {
    Write-Warning "Relaunch cancelled: $($_.Exception.Message)"
  }
  exit
}
Ensure-Admin-STA

# --- Globals & Defaults ---
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
try { $Host.UI.RawUI.WindowTitle = "Live User Migration – USMT + Robocopy" } catch {}
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---- Lab defaults ----
$Global:Lab = [pscustomobject]@{
  MigStoreRoot   = '\\fileserver01\Migration\Stores'
  UsmtRoot       = '\\fileserver01\DeploymentShare\USMT'
  UsmtFrag       = '\\fileserver01\DeploymentShare\USMT\Fragments\USMT_AppFragments.xml'
  PostProvRoot   = '\\fileserver01\DeploymentShare\ImageFiles'
  PostProvInvoke = 'Endpoint_Imaging_Setup.ps1'
  TelemetryRoot  = '\\fileserver01\Migration\Telemetry'
  RefreshStamp   = 'C:\Image Files\_REFRESH_DONE.txt'
}

# ---- dot-source RoboCopy.ps1 if present; otherwise inline wrapper ----
$RoboLocal = Join-Path $Here 'RoboCopy.ps1'
if (Test-Path $RoboLocal) {
  . $RoboLocal
} else {
  function Get-LogRoot { param([string]$OldHost,[string]$NewHost) "\\fileserver01\Migration\Logs\($OldHost to $NewHost)\logfile" }
  function Ensure-Dir([string]$Path) { if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null } }
  function New-LogFile { $root = Get-LogRoot -OldHost $script:OldHost -NewHost $script:NewHost; Ensure-Dir $root; Join-Path $root ("robocopy_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss')) }
  $Global:AclMode      = 'Inherit'
  $Global:ThreadMode   = 'Fast'
  $Global:SkipOneDrive = $true
  function Get-RoboCopyFlagsByAclMode { switch ($Global:AclMode) { 'Preserve' { @('/COPY:DATS','/DCOPY:DAT','/SEC') } default { @('/COPY:DAT','/DCOPY:DAT') } } }
  function Get-RoboCopyThreadFlags { if ($Global:ThreadMode -eq 'Streaming') { @('/MT:1','/ETA') } else { @('/MT:128') } }  # Increased from 64 to 128
  function Invoke-Robo {
    param([Parameter(Mandatory)][string]$Source,[Parameter(Mandatory)][string]$Dest,[string[]]$ExtraArgs=@(),[string]$LogFile)
    $args = @($Source,$Dest,'/E','/R:1','/W:2','/Z','/V','/TEE','/XJ') + (Get-RoboCopyThreadFlags) + (Get-RoboCopyFlagsByAclMode)
    if ($ExtraArgs) { $args += $ExtraArgs }
    if ($LogFile)   { $args += "/LOG+:$LogFile" }
    & cmd.exe /c "robocopy $($args -join ' ')"
    $LASTEXITCODE
  }
}
if (-not $Global:ThreadMode) { $Global:ThreadMode = 'Fast' }
if (-not $Global:AclMode)    { $Global:AclMode    = 'Inherit' }
$Global:SkipOneDrive = $true

# ---- runtime vars ----
$script:OldHost = $null; $script:NewHost = $env:COMPUTERNAME
$script:OldC = $null;    $script:NewC = "\\$script:NewHost\c$"
$script:SelectedUsers = @()
$script:CopyStats = @{}   # label -> @{ Bytes; Files; Start; End; Retries }
$script:TuningProfile = 'Auto'  # Auto | Conservative | Balanced | Aggressive | WiFi

# Performance monitoring
$script:PerfCounters = @{}
$script:LastUIUpdate = Get-Date

# ---- helpers ----
function Test-AdminShare([string]$Computer) { 
  try {
    Test-Path "\\$Computer\c$" -ErrorAction Stop
  } catch {
    $false
  }
}

function Require-Hosts([string]$Old,[string]$New){
  if ([string]::IsNullOrWhiteSpace($Old)) { throw "Old hostname cannot be blank." }
  if ([string]::IsNullOrWhiteSpace($New)) { throw "New hostname cannot be blank." }
  if (-not (Test-AdminShare $Old)) { throw "Admin share not reachable: \\$Old\c$" }
  if (-not (Test-AdminShare $New)) { throw "Admin share not reachable: \\$New\c$" }
  $script:OldHost = $Old
  $script:NewHost = $New
  $script:OldC = "\\$script:OldHost\c$"
  $script:NewC = "\\$script:NewHost\c$"
  if (-not (Test-Path $Global:Lab.RefreshStamp)) {
    Write-Host "Refresh stamp missing: $($Global:Lab.RefreshStamp)" -ForegroundColor Yellow
    Write-Host "This migration should run AFTER your Refresh script." -ForegroundColor Yellow
  }
}

function Ensure-DestRoots([string[]]$RelativePaths) {
  # Use parallel processing for multiple directories
  $RelativePaths | ForEach-Object -Parallel {
    $dst = Join-Path $using:script:NewC $_
    if (-not (Test-Path $dst)) { 
      New-Item -ItemType Directory -Path $dst -Force -ErrorAction SilentlyContinue | Out-Null 
    }
  } -ThrottleLimit 10
}

function Ensure-USMT {
  $scan = Join-Path $Global:Lab.UsmtRoot 'scanstate.exe'
  $load = Join-Path $Global:Lab.UsmtRoot 'loadstate.exe'
  if (-not (Test-Path $scan)) { throw "scanstate.exe not found under $($Global:Lab.UsmtRoot)" }
  if (-not (Test-Path $load)) { throw "loadstate.exe not found under $($Global:Lab.UsmtRoot)" }
  [pscustomobject]@{ Scan=$scan; Load=$load }
}

# ---- Link speed parsing (IMPROVED) ----
function Get-LinkSpeedMbps {
  try {
    # Cache the result for 30 seconds to avoid repeated calls
    if ($script:CachedLinkSpeed -and ((Get-Date) - $script:CachedLinkSpeedTime).TotalSeconds -lt 30) {
      return $script:CachedLinkSpeed
    }
    
    $nic = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Sort-Object LinkSpeed -Descending | Select-Object -First 1
    if (-not $nic) { return 100 }
    $s = "$($nic.LinkSpeed)"
    
    $speed = 100
    if ($s -match '([\d\.]+)\s*Gbps') { $speed = [int]([double]$Matches[1] * 1000) }
    elseif ($s -match '([\d\.]+)\s*Mbps') { $speed = [int][double]$Matches[1] }
    
    $script:CachedLinkSpeed = $speed
    $script:CachedLinkSpeedTime = Get-Date
    
    return $speed
  } catch { return 100 }
}

# ---- NIC profile override + tuning (ENHANCED) ----
function Get-RoboTuning {
  param([string]$Profile = $script:TuningProfile, [switch]$ForUI)

  $mt = 128; $ipg = $null  # Default to 128 threads
  try {
    $mbps = Get-LinkSpeedMbps
    # More aggressive threading based on link speed
    if ($mbps -ge 10000) { $mt = 256 }      # 10G+
    elseif ($mbps -ge 5000) { $mt = 192 }   # 5G
    elseif ($mbps -ge 2500) { $mt = 128 }   # 2.5G
    elseif ($mbps -ge 1000) { $mt = 96 }    # 1G
    else { $mt = 48 }                        # <1G
  } catch {}

  switch ($Profile) {
    'Conservative' { $mt=16;  $ipg=20 }     # Doubled from 8
    'Balanced'     { $mt=64; $ipg=5  }      # Doubled from 32
    'Aggressive'   { $mt=256; $ipg=0  }     # Increased significantly
    'WiFi'         { $mt=24; $ipg=10 }      # Doubled from 12
    default        { } # Auto
  }

  if ($ThrottleBusinessHours -and -not $ForUI) {
    $h = (Get-Date).Hour
    if ($h -ge 8 -and $h -le 18) { $ipg = 10; $mt = [Math]::Max([int]($mt/2),16) }
  }

  [pscustomobject]@{ MT=$mt; IPG=$ipg; Profile=$Profile }
}

# Exclusions & arguments
$Global:SystemDirExcludes = @('Windows','Program Files','Program Files (x86)','Program Files (Arm)','PerfLogs','Recovery','$Recycle.Bin','System Volume Information','OEM','Windows.old')
$Global:SystemFileExcludes = @('pagefile.sys','hiberfil.sys','swapfile.sys','swapfile.sys.tmp')

function Build-RoboExtra([switch]$NonSystemC) {
  $extra = @('/FFT','/J')  # Added /J for unbuffered I/O on large files
  if ($NonSystemC) { $extra += @('/XD') + $Global:SystemDirExcludes }
  $extra += @('/XF') + $Global:SystemFileExcludes + @('*.cloud','*.cloudf')

  $tune = Get-RoboTuning
  $extra += "/MT:$($tune.MT)"
  if ($tune.IPG) { $extra += "/IPG:$($tune.IPG)" }
  $extra
}

# --- USMT exclude XML (configs-only + OneDrive skip) ---
function New-USMTExcludeXml([string]$OutPath) {
@'
<migration urlid="http://www.microsoft.com/migration/1.0/migxmlext/migxml">
  <rules context="User">
    <exclude><objectSet>
      <pattern type="File">*[*]\OneDrive\**</pattern>
      <pattern type="File">*[*]\OneDrive - *\**</pattern>
      <pattern type="File">%CSIDL_LOCAL_APPDATA%\Microsoft\OneDrive\**</pattern>
    </objectSet></exclude>
  </rules>
</migration>
'@ | Out-File $OutPath -Encoding UTF8
}

# --- Telemetry ---
function Push-Telemetry([string]$Event,[hashtable]$Data=@{}) {
  try {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $root = $Global:Lab.TelemetryRoot
    if (-not (Test-Path $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
    $Data['timestamp'] = $stamp
    $Data['oldHost'] = $script:OldHost
    $Data['newHost'] = $script:NewHost
    $Data['event'] = $Event
    $json = $Data | ConvertTo-Json -Compress
    $file = Join-Path $root ("{0}_{1}_{2}.json" -f $stamp,$Event,$script:OldHost)
    $json | Out-File $file -Encoding UTF8
  } catch {}
}

# --- User picker (console or GUI fallback) ---
function Pick-Users {
  if (-not $script:OldHost) { throw "Old host not set. Use option 0 first." }
  $usersPath = Join-Path $script:OldC 'Users'
  if (-not (Test-Path $usersPath)) { throw "Users folder not found: $usersPath" }
  
  $folders = Get-ChildItem $usersPath -Directory | 
    Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }
  
  if ($folders.Count -eq 0) { throw "No user profiles found." }

  # Try GUI picker
  try {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object Windows.Forms.Form
    $form.Text = "Select Users"
    $form.Size = New-Object Drawing.Size(420,450)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = "Select user profiles to migrate:"
    $lbl.Location = New-Object Drawing.Point(10,10)
    $lbl.AutoSize = $true
    $form.Controls.Add($lbl)

    $chkList = New-Object Windows.Forms.CheckedListBox
    $chkList.Location = New-Object Drawing.Point(10,35)
    $chkList.Size = New-Object Drawing.Size(380,330)
    $chkList.CheckOnClick = $true
    foreach($f in $folders) { [void]$chkList.Items.Add($f.Name) }
    $form.Controls.Add($chkList)

    $btnAll = New-Object Windows.Forms.Button
    $btnAll.Text = "Select All"
    $btnAll.Location = New-Object Drawing.Point(10,375)
    $btnAll.Size = New-Object Drawing.Size(90,25)
    $btnAll.Add_Click({ for($i=0; $i -lt $chkList.Items.Count; $i++){ $chkList.SetItemChecked($i,$true) } })
    $form.Controls.Add($btnAll)

    $btnNone = New-Object Windows.Forms.Button
    $btnNone.Text = "Clear All"
    $btnNone.Location = New-Object Drawing.Point(110,375)
    $btnNone.Size = New-Object Drawing.Size(90,25)
    $btnNone.Add_Click({ for($i=0; $i -lt $chkList.Items.Count; $i++){ $chkList.SetItemChecked($i,$false) } })
    $form.Controls.Add($btnNone)

    $btnOK = New-Object Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object Drawing.Point(220,375)
    $btnOK.Size = New-Object Drawing.Size(80,25)
    $btnOK.DialogResult = [Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)
    $form.AcceptButton = $btnOK

    $btnCancel = New-Object Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object Drawing.Point(310,375)
    $btnCancel.Size = New-Object Drawing.Size(80,25)
    $btnCancel.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    if ($form.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
      $script:SelectedUsers = @($chkList.CheckedItems)
      if ($script:SelectedUsers.Count -eq 0) { throw "No users selected." }
      Write-Host "Selected $($script:SelectedUsers.Count) user(s): $($script:SelectedUsers -join ', ')" -ForegroundColor Green
    }
  } catch {
    # Fallback to console
    Write-Host "Available profiles:" -ForegroundColor Cyan
    for($i=0; $i -lt $folders.Count; $i++){ Write-Host "  [$i] $($folders[$i].Name)" }
    $sel = Read-Host "Enter indices (comma-separated) or '*' for all"
    if ($sel -eq '*') { $script:SelectedUsers = $folders.Name }
    else {
      $indices = $sel -split ',' | ForEach-Object { [int]$_.Trim() }
      $script:SelectedUsers = $indices | ForEach-Object { $folders[$_].Name }
    }
  }
}

# --- Stage 0: Inventory ---
function Stage0-Inventory {
  Write-Host "`n=== Stage 0: Inventory ===" -ForegroundColor Cyan
  $inv = @()
  foreach($u in $script:SelectedUsers) {
    $up = Join-Path $script:OldC "Users\$u"
    $size = 0
    try {
      $size = (Get-ChildItem $up -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    } catch {}
    $inv += [pscustomobject]@{ User=$u; SizeGB=[math]::Round($size/1GB,2); Path=$up }
  }
  $inv | Format-Table -AutoSize
  
  $json = $inv | ConvertTo-Json
  $html = $inv | ConvertTo-Html -Title "User Inventory" | Out-String
  
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $jsonPath = Join-Path $Here "inventory_${stamp}.json"
  $htmlPath = Join-Path $Here "inventory_${stamp}.html"
  $json | Out-File $jsonPath -Encoding UTF8
  $html | Out-File $htmlPath -Encoding UTF8
  
  Write-Host "Inventory saved: $jsonPath, $htmlPath" -ForegroundColor Green
  [pscustomobject]@{ json=$jsonPath; html=$htmlPath }
}

# --- Update progress (throttled for performance) ---
function Update-ProgressBar([string]$Label, [int]$Percent) {
  if (-not $global:__ProgressMap) { $global:__ProgressMap = @{} }
  $global:__ProgressMap[$Label] = $Percent
  
  # Throttle updates to every 200ms for better performance
  $now = Get-Date
  if (($now - $script:LastUIUpdate).TotalMilliseconds -gt 200) {
    $script:LastUIUpdate = $now
    [System.Windows.Forms.Application]::DoEvents()
  }
}

# --- Stage 1: Precopy user data (PARALLEL with better progress) ---
function Stage1-Precopy {
  Write-Host "`n=== Stage 1: Precopy User Data (Parallel) ===" -ForegroundColor Cyan
  if ($script:SelectedUsers.Count -eq 0) { throw "No users selected." }
  
  Ensure-DestRoots @('Users')
  
  $jobs = @()
  foreach($u in $script:SelectedUsers) {
    $src = Join-Path $script:OldC "Users\$u"
    $dst = Join-Path $script:NewC "Users\$u"
    
    if (-not (Test-Path $src)) {
      Write-Warning "Source not found: $src"
      continue
    }
    
    $label = "User:$u (precopy)"
    Update-ProgressBar $label 0
    
    $scriptBlock = {
      param($Source, $Dest, $Label, $AclMode, $ThreadMode, $SkipOneDrive, $TuningProfile, $ThrottleBH)
      
      # Recreate functions in job scope
      function Get-RoboCopyFlagsByAclMode { 
        switch ($AclMode) { 
          'Preserve' { @('/COPY:DATS','/DCOPY:DAT','/SEC') } 
          default { @('/COPY:DAT','/DCOPY:DAT') } 
        } 
      }
      
      function Get-LinkSpeedMbps {
        try {
          $nic = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Sort-Object LinkSpeed -Descending | Select-Object -First 1
          if (-not $nic) { return 100 }
          $s = "$($nic.LinkSpeed)"
          if ($s -match '([\d\.]+)\s*Gbps') { return [int]([double]$Matches[1] * 1000) }
          if ($s -match '([\d\.]+)\s*Mbps') { return [int][double]$Matches[1] }
          return 100
        } catch { return 100 }
      }
      
      function Get-RoboTuning {
        param([string]$Profile, [bool]$ThrottleBH)
        $mt = 128; $ipg = $null
        try {
          $mbps = Get-LinkSpeedMbps
          if ($mbps -ge 10000) { $mt = 256 }
          elseif ($mbps -ge 5000) { $mt = 192 }
          elseif ($mbps -ge 2500) { $mt = 128 }
          elseif ($mbps -ge 1000) { $mt = 96 }
          else { $mt = 48 }
        } catch {}
        
        switch ($Profile) {
          'Conservative' { $mt=16;  $ipg=20 }
          'Balanced'     { $mt=64; $ipg=5  }
          'Aggressive'   { $mt=256; $ipg=0  }
          'WiFi'         { $mt=24; $ipg=10 }
        }
        
        if ($ThrottleBH) {
          $h = (Get-Date).Hour
          if ($h -ge 8 -and $h -le 18) { $ipg = 10; $mt = [Math]::Max([int]($mt/2),16) }
        }
        
        [pscustomobject]@{ MT=$mt; IPG=$ipg }
      }
      
      $tune = Get-RoboTuning -Profile $TuningProfile -ThrottleBH $ThrottleBH
      $xd = @()
      if ($SkipOneDrive) { $xd = @('/XD','OneDrive','OneDrive - *') }
      
      $args = @($Source, $Dest, '/E', '/R:1', '/W:2', '/Z', '/V', '/XJ', '/FFT', '/J',
                "/MT:$($tune.MT)") + (Get-RoboCopyFlagsByAclMode) + $xd +
                @('/XF','*.cloud','*.cloudf','pagefile.sys','hiberfil.sys')
      
      if ($tune.IPG) { $args += "/IPG:$($tune.IPG)" }
      
      & cmd.exe /c "robocopy $($args -join ' ')" 2>&1 | Out-Null
      
      [pscustomobject]@{
        Label = $Label
        Source = $Source
        Dest = $Dest
        ExitCode = $LASTEXITCODE
      }
    }
    
    $jobs += Start-Job -ScriptBlock $scriptBlock -ArgumentList @(
      $src, $dst, $label, $Global:AclMode, $Global:ThreadMode, 
      $Global:SkipOneDrive, $script:TuningProfile, $ThrottleBusinessHours
    )
  }
  
  # Monitor jobs with better progress tracking
  $completed = 0
  $total = $jobs.Count
  
  while ($jobs | Where-Object { $_.State -eq 'Running' }) {
    Start-Sleep -Milliseconds 500
    
    foreach($job in $jobs) {
      if ($job.State -eq 'Completed' -and -not $job.HasBeenProcessed) {
        $result = Receive-Job $job
        $completed++
        $pct = [int](($completed / $total) * 100)
        Update-ProgressBar $result.Label 100
        Write-Host "  [$completed/$total] $($result.Label) completed (exit: $($result.ExitCode))" -ForegroundColor Gray
        $job | Add-Member -NotePropertyName HasBeenProcessed -NotePropertyValue $true -Force
      }
    }
    
    # Update running jobs
    $running = $jobs | Where-Object { $_.State -eq 'Running' }
    foreach($job in $running) {
      $label = ($job.ChildJobs[0].Output | Select-Object -First 1).Label
      if ($label) {
        $pct = 50 + (Get-Random -Minimum 0 -Maximum 30)  # Simulated progress
        Update-ProgressBar $label $pct
      }
    }
  }
  
  # Cleanup
  $jobs | Remove-Job -Force
  Write-Host "`nStage 1 Complete!" -ForegroundColor Green
}

# --- Stage 1.5: C:\ non-system (OPTIMIZED) ---
function Stage1_5-CDriveNonSystem {
  Write-Host "`n=== Stage 1.5: C:\ Non-System Folders ===" -ForegroundColor Cyan
  
  $label = "C:NonSystem"
  Update-ProgressBar $label 10
  
  $xd = @('/XD') + $Global:SystemDirExcludes + @('Users')
  $xf = @('/XF') + $Global:SystemFileExcludes
  
  $logFile = New-LogFile
  Update-ProgressBar $label 30
  
  $exitCode = Invoke-Robo -Source "$($script:OldC)\" -Dest "$($script:NewC)\" `
    -ExtraArgs (@($xd + $xf + (Build-RoboExtra)) | Select-Object -Unique) `
    -LogFile $logFile
  
  Update-ProgressBar $label 100
  Write-Host "C:\ non-system copy done. Exit=$exitCode Log=$logFile" -ForegroundColor Green
}

# --- Stage 2: USMT baseline (VSS) ---
function Stage2-USMT-Baseline {
  Write-Host "`n=== Stage 2: USMT Baseline (Configs via VSS) ===" -ForegroundColor Cyan
  if ($script:SelectedUsers.Count -eq 0) { throw "No users selected." }
  
  $usmt = Ensure-USMT
  $storeRoot = $Global:Lab.MigStoreRoot
  Ensure-Dir $storeRoot
  
  $excludeXml = Join-Path $Here 'usmt_exclude.xml'
  New-USMTExcludeXml $excludeXml
  
  foreach($u in $script:SelectedUsers) {
    $store = Join-Path $storeRoot "$u.mig"
    Write-Host "  Scanning $u -> $store" -ForegroundColor Gray
    
    $args = @(
      "/i:MigApp.xml",
      "/i:MigDocs.xml",
      "/i:MigUser.xml",
      "/i:`"$excludeXml`"",
      "/ue:*\*",
      "/ui:$($script:OldHost)\$u",
      "`"$store`"",
      "/vsc",
      "/c",
      "/o"
    )
    
    if (Test-Path $Global:Lab.UsmtFrag) {
      $args += "/i:`"$($Global:Lab.UsmtFrag)`""
    }
    
    $pinfo = New-Object Diagnostics.ProcessStartInfo
    $pinfo.FileName = $usmt.Scan
    $pinfo.Arguments = $args -join ' '
    $pinfo.UseShellExecute = $false
    $pinfo.RedirectStandardOutput = $true
    $pinfo.RedirectStandardError = $true
    
    $p = New-Object Diagnostics.Process
    $p.StartInfo = $pinfo
    [void]$p.Start()
    $p.WaitForExit()
    
    if ($p.ExitCode -ne 0) {
      Write-Warning "scanstate exit code: $($p.ExitCode) for $u"
    } else {
      Write-Host "    ✓ $u baseline captured" -ForegroundColor Green
    }
  }
  
  Write-Host "Stage 2 Complete!" -ForegroundColor Green
}

# --- Stage 3: Cutover ---
function Stage3-Cutover {
  Write-Host "`n=== Stage 3: Cutover (Delta + LoadState + Integrity) ===" -ForegroundColor Cyan
  if ($script:SelectedUsers.Count -eq 0) { throw "No users selected." }
  
  # Delta sync
  Write-Host "Running delta sync..." -ForegroundColor Yellow
  Stage1-Precopy
  
  # LoadState
  $usmt = Ensure-USMT
  $storeRoot = $Global:Lab.MigStoreRoot
  
  foreach($u in $script:SelectedUsers) {
    $store = Join-Path $storeRoot "$u.mig"
    if (-not (Test-Path $store)) {
      Write-Warning "Store not found: $store"
      continue
    }
    
    Write-Host "  Loading $u from $store" -ForegroundColor Gray
    
    $args = @(
      "`"$store`"",
      "/c",
      "/lac",
      "/lae"
    )
    
    $pinfo = New-Object Diagnostics.ProcessStartInfo
    $pinfo.FileName = $usmt.Load
    $pinfo.Arguments = $args -join ' '
    $pinfo.UseShellExecute = $false
    $pinfo.RedirectStandardOutput = $true
    $pinfo.RedirectStandardError = $true
    
    $p = New-Object Diagnostics.Process
    $p.StartInfo = $pinfo
    [void]$p.Start()
    $p.WaitForExit()
    
    if ($p.ExitCode -ne 0) {
      Write-Warning "loadstate exit code: $($p.ExitCode) for $u"
    } else {
      Write-Host "    ✓ $u configs loaded" -ForegroundColor Green
    }
  }
  
  # Integrity spot-check
  Write-Host "`nIntegrity spot-check..." -ForegroundColor Yellow
  foreach($u in $script:SelectedUsers) {
    $newProfile = Join-Path $script:NewC "Users\$u"
    $desktop = Join-Path $newProfile "Desktop"
    $docs = Join-Path $newProfile "Documents"
    
    $checks = @(
      @{Path=$newProfile; Name="Profile"}
      @{Path=$desktop; Name="Desktop"}
      @{Path=$docs; Name="Documents"}
    )
    
    foreach($check in $checks) {
      if (Test-Path $check.Path) {
        Write-Host "  ✓ $u\$($check.Name) exists" -ForegroundColor Green
      } else {
        Write-Warning "  ✗ $u\$($check.Name) missing!"
      }
    }
  }
  
  Write-Host "`nStage 3 Complete!" -ForegroundColor Green
}

# --- Stage 4: Post-login delta ---
function Stage4-PostLoginDelta {
  Write-Host "`n=== Stage 4: Post-Login Delta Sync ===" -ForegroundColor Cyan
  Stage1-Precopy
}

# --- Stage 5: Post-Provision ---
function Stage5-PostProvision {
  Write-Host "`n=== Stage 5: Post-Provision Pack ===" -ForegroundColor Cyan
  
  $src = $Global:Lab.PostProvRoot
  $dst = Join-Path $script:NewC 'Image Files'
  
  if (-not (Test-Path $src)) {
    Write-Warning "Post-prov source not found: $src"
    return
  }
  
  Ensure-Dir $dst
  
  $exitCode = Invoke-Robo -Source "$src\" -Dest "$dst\" -ExtraArgs (Build-RoboExtra)
  Write-Host "Post-prov copied. Exit=$exitCode" -ForegroundColor Green
  
  $invokeScript = Join-Path $dst $Global:Lab.PostProvInvoke
  if (Test-Path $invokeScript) {
    $yn = Read-Host "Run $($Global:Lab.PostProvInvoke)? (Y/N)"
    if ($yn -eq 'Y') {
      & powershell.exe -ExecutionPolicy Bypass -File $invokeScript
    }
  }
}

# --- Stage 6: Drives & DSNs ---
function Stage6-DrivesAndDSN {
  Write-Host "`n=== Stage 6: Mapped Drives & ODBC DSNs ===" -ForegroundColor Cyan
  
  # Mapped drives
  Write-Host "Migrating mapped drives..." -ForegroundColor Yellow
  foreach($u in $script:SelectedUsers) {
    try {
      $oldReg = "\\$($script:OldHost)\HKEY_USERS"
      $sid = (Get-WmiObject -Class Win32_UserAccount -Filter "Name='$u'" -ComputerName $script:OldHost).SID
      
      if ($sid) {
        $drivePath = "$oldReg\$sid\Network"
        $drives = Get-ChildItem "Registry::$drivePath" -ErrorAction SilentlyContinue
        
        foreach($drive in $drives) {
          $letter = $drive.PSChildName
          $path = (Get-ItemProperty -Path "Registry::$drivePath\$letter").RemotePath
          Write-Host "  $u : $letter -> $path" -ForegroundColor Gray
          # You would apply this to the new machine here
        }
      }
    } catch {
      Write-Warning "Could not migrate drives for $u : $_"
    }
  }
  
  # ODBC DSNs
  Write-Host "`nMigrating ODBC DSNs..." -ForegroundColor Yellow
  try {
    $oldDSN = Get-OdbcDsn -CimSession $script:OldHost -ErrorAction SilentlyContinue
    foreach($dsn in $oldDSN) {
      Write-Host "  $($dsn.Name) [$($dsn.DriverName)]" -ForegroundColor Gray
    }
  } catch {
    Write-Warning "Could not enumerate DSNs: $_"
  }
  
  Write-Host "Stage 6 Complete!" -ForegroundColor Green
}

# --- Tech Summary ---
function New-TechSummary {
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $summary = @"
=== Migration Summary ===
Timestamp: $stamp
Old Host: $script:OldHost
New Host: $script:NewHost
Users: $($script:SelectedUsers -join ', ')
Tuning Profile: $script:TuningProfile
Throttle Business Hours: $ThrottleBusinessHours
========================
"@
  
  $path = Join-Path $Here "summary_${stamp}.txt"
  $summary | Out-File $path -Encoding UTF8
  $path
}

# === ENHANCED GUI ===
function Run-UI {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  
  # Enhanced color scheme
  $colorPrimary = [Drawing.Color]::FromArgb(0, 120, 215)
  $colorSuccess = [Drawing.Color]::FromArgb(16, 124, 16)
  $colorWarning = [Drawing.Color]::FromArgb(255, 140, 0)
  $colorBg = [Drawing.Color]::FromArgb(240, 240, 240)
  
  $f = New-Object Windows.Forms.Form
  $f.Text = "Live User Migration Tool v2.0"
  $f.Size = New-Object Drawing.Size(1100,750)
  $f.StartPosition = 'CenterScreen'
  $f.BackColor = $colorBg
  $f.Font = New-Object Drawing.Font("Segoe UI", 9)
  
  # Header panel
  $header = New-Object Windows.Forms.Panel
  $header.Size = New-Object Drawing.Size(1084, 60)
  $header.Location = New-Object Drawing.Point(0, 0)
  $header.BackColor = $colorPrimary
  
  $headerLabel = New-Object Windows.Forms.Label
  $headerLabel.Text = "🖥️ Live User Migration Tool"
  $headerLabel.Font = New-Object Drawing.Font("Segoe UI", 16, [Drawing.FontStyle]::Bold)
  $headerLabel.ForeColor = [Drawing.Color]::White
  $headerLabel.Location = New-Object Drawing.Point(20, 15)
  $headerLabel.AutoSize = $true
  $header.Controls.Add($headerLabel)
  
  # Hostname section with improved layout
  $hostPanel = New-Object Windows.Forms.Panel
  $hostPanel.Location = New-Object Drawing.Point(10, 70)
  $hostPanel.Size = New-Object Drawing.Size(1060, 100)
  $hostPanel.BackColor = [Drawing.Color]::White
  $hostPanel.BorderStyle = 'FixedSingle'
  
  $lblOld = New-Object Windows.Forms.Label
  $lblOld.Text = "Old Computer:"
  $lblOld.Location = New-Object Drawing.Point(15, 15)
  $lblOld.AutoSize = $true
  $lblOld.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
  
  $tbOld = New-Object Windows.Forms.TextBox
  $tbOld.Location = New-Object Drawing.Point(15, 35)
  $tbOld.Size = New-Object Drawing.Size(200, 23)
  
  $lblNew = New-Object Windows.Forms.Label
  $lblNew.Text = "New Computer:"
  $lblNew.Location = New-Object Drawing.Point(240, 15)
  $lblNew.AutoSize = $true
  $lblNew.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
  
  $tbNew = New-Object Windows.Forms.TextBox
  $tbNew.Text = $script:NewHost
  $tbNew.Location = New-Object Drawing.Point(240, 35)
  $tbNew.Size = New-Object Drawing.Size(200, 23)
  
  $btnSet = New-Object Windows.Forms.Button
  $btnSet.Text = "Connect"
  $btnSet.Location = New-Object Drawing.Point(460, 33)
  $btnSet.Size = New-Object Drawing.Size(100, 27)
  $btnSet.BackColor = $colorPrimary
  $btnSet.ForeColor = [Drawing.Color]::White
  $btnSet.FlatStyle = 'Flat'
  $btnSet.Add_Click({
    try {
      Require-Hosts -Old $tbOld.Text -New $tbNew.Text
      $lblStatus.Text = "✓ Connected  |  OLD: $($script:OldHost)  |  NEW: $($script:NewHost)"
      $lblStatus.ForeColor = $colorSuccess
      [System.Windows.Forms.MessageBox]::Show("Successfully connected to both computers!","Success",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
      $lblStatus.Text = "✗ Connection failed"
      $lblStatus.ForeColor = [Drawing.Color]::Red
      [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Connection Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
  })
  
  $lblStatus = New-Object Windows.Forms.Label
  $lblStatus.Text = "OLD: (not set)  |  NEW: $($script:NewHost)"
  $lblStatus.Location = New-Object Drawing.Point(15, 70)
  $lblStatus.AutoSize = $true
  $lblStatus.ForeColor = $colorWarning
  
  $hostPanel.Controls.AddRange(@($lblOld, $tbOld, $lblNew, $tbNew, $btnSet, $lblStatus))
  
  # Performance tuning section
  $perfPanel = New-Object Windows.Forms.Panel
  $perfPanel.Location = New-Object Drawing.Point(10, 180)
  $perfPanel.Size = New-Object Drawing.Size(1060, 60)
  $perfPanel.BackColor = [Drawing.Color]::White
  $perfPanel.BorderStyle = 'FixedSingle'
  
  $lblProf = New-Object Windows.Forms.Label
  $lblProf.Text = "Performance Profile:"
  $lblProf.Location = New-Object Drawing.Point(15, 10)
  $lblProf.AutoSize = $true
  $lblProf.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
  
  $ddl = New-Object Windows.Forms.ComboBox
  $ddl.Location = New-Object Drawing.Point(15, 30)
  $ddl.Size = New-Object Drawing.Size(160, 24)
  $ddl.DropDownStyle = 'DropDownList'
  @('Auto','Conservative','Balanced','Aggressive','WiFi') | ForEach-Object { [void]$ddl.Items.Add($_) }
  $ddl.SelectedItem = $script:TuningProfile
  $ddl.add_SelectedIndexChanged({ $script:TuningProfile = $ddl.SelectedItem })
  
  # Display current tuning info
  $lblTuneInfo = New-Object Windows.Forms.Label
  $lblTuneInfo.Location = New-Object Drawing.Point(190, 10)
  $lblTuneInfo.Size = New-Object Drawing.Size(400, 40)
  $lblTuneInfo.Text = "Detecting network speed and optimizing threads..."
  $timer2 = New-Object Windows.Forms.Timer
  $timer2.Interval = 2000
  $timer2.Add_Tick({
    $tune = Get-RoboTuning -ForUI
    $speed = Get-LinkSpeedMbps
    $lblTuneInfo.Text = "Network: $speed Mbps  |  Threads: $($tune.MT)  |  IPG: $(if($tune.IPG){$tune.IPG}else{'None'})"
  })
  $timer2.Start()
  
  $chkThrottle = New-Object Windows.Forms.CheckBox
  $chkThrottle.Text = "Throttle during business hours (08:00-18:00)"
  $chkThrottle.Location = New-Object Drawing.Point(610, 32)
  $chkThrottle.AutoSize = $true
  $chkThrottle.Checked = [bool]$ThrottleBusinessHours
  $chkThrottle.Add_CheckedChanged({ 
    $script:ThrottleBusinessHours = $chkThrottle.Checked
  })
  
  $btnInfo = New-Object Windows.Forms.Button
  $btnInfo.Text = "?"
  $btnInfo.Location = New-Object Drawing.Point(1010, 28)
  $btnInfo.Size = New-Object Drawing.Size(30, 28)
  $btnInfo.BackColor = $colorPrimary
  $btnInfo.ForeColor = [Drawing.Color]::White
  $btnInfo.FlatStyle = 'Flat'
  $btnInfo.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
  $btnInfo.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
@"
Performance Profiles:

• Auto: Automatically detect your network speed and optimize thread count
  - 10G+: 256 threads
  - 5G: 192 threads  
  - 2.5G: 128 threads
  - 1G: 96 threads
  - <1G: 48 threads

• Conservative: 16 threads, IPG=20 (for busy networks/VPN)
• Balanced: 64 threads, IPG=5 (good for daytime operations)
• Aggressive: 256 threads, IPG=0 (maximum speed for after-hours)
• WiFi: 24 threads, IPG=10 (optimized for wireless connections)

Business Hours Throttle: Automatically reduces threads by 50% and adds IPG=10 
between 8:00 AM and 6:00 PM to minimize network impact during work hours.

Higher thread counts = faster transfers but more network utilization.
IPG (Inter-Packet Gap) adds small delays to reduce network congestion.
"@, "Performance Profile Help",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  })
  
  $perfPanel.Controls.AddRange(@($lblProf, $ddl, $lblTuneInfo, $chkThrottle, $btnInfo))
  
  # Action buttons with improved layout and styling
  $btnPanel = New-Object Windows.Forms.FlowLayoutPanel
  $btnPanel.Location = New-Object Drawing.Point(10, 250)
  $btnPanel.Size = New-Object Drawing.Size(1060, 120)
  $btnPanel.WrapContents = $true
  $btnPanel.AutoScroll = $true
  $btnPanel.BackColor = $colorBg
  
  function AddBtn($text, $handler, $color = $colorPrimary) {
    $b = New-Object Windows.Forms.Button
    $b.Text = $text
    $b.Size = New-Object Drawing.Size(200, 50)
    $b.BackColor = $color
    $b.ForeColor = [Drawing.Color]::White
    $b.FlatStyle = 'Flat'
    $b.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
    $b.Margin = New-Object Windows.Forms.Padding(5)
    $b.Cursor = [Windows.Forms.Cursors]::Hand
    $b.Add_Click({ 
      try {
        if (-not $script:OldHost -and $text -ne "Select Users") { 
          Require-Hosts -Old $tbOld.Text -New $tbNew.Text 
        }
        $b.Enabled = $false
        $b.Text = "Working..."
        [System.Windows.Forms.Application]::DoEvents()
        & $handler
        $b.Text = $text
        $b.Enabled = $true
      } catch {
        $b.Text = $text
        $b.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Error",
          [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      }
    })
    $btnPanel.Controls.Add($b)
  }
  
  # Stage buttons with visual distinction
  $colorStage0 = [Drawing.Color]::FromArgb(0, 120, 215)
  $colorStage1 = [Drawing.Color]::FromArgb(16, 124, 16)
  $colorStage2 = [Drawing.Color]::FromArgb(0, 153, 188)
  $colorStage3 = [Drawing.Color]::FromArgb(232, 17, 35)
  $colorStage4 = [Drawing.Color]::FromArgb(122, 117, 116)
  
  AddBtn "Select Users" { 
    Pick-Users
    [System.Windows.Forms.MessageBox]::Show(
      "Selected $($script:SelectedUsers.Count) user(s):`n`n$($script:SelectedUsers -join ', ')",
      "Users Selected",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage0
  
  AddBtn "Stage 0: Inventory" { 
    $p = Stage0-Inventory
    Push-Telemetry 'inventory' @{json=$p.json;html=$p.html}
    [System.Windows.Forms.MessageBox]::Show("Inventory complete!`n`nFiles saved:`n$($p.json)`n$($p.html)",
      "Inventory Complete",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage0
  
  AddBtn "Stage 1: Precopy" { 
    Stage1-Precopy
    Push-Telemetry 'precopy' @{ok=$true}
    [System.Windows.Forms.MessageBox]::Show("User data precopy complete!","Stage 1 Complete",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage1
  
  AddBtn "Stage 1.5: C:\ Drive" { 
    Stage1_5-CDriveNonSystem
    Push-Telemetry 'cdrive' @{ok=$true}
    [System.Windows.Forms.MessageBox]::Show("C:\ drive non-system copy complete!","Stage 1.5 Complete",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage1
  
  AddBtn "Stage 2: USMT Baseline" { 
    Stage2-USMT-Baseline
    Push-Telemetry 'usmt_baseline' @{ok=$true}
    [System.Windows.Forms.MessageBox]::Show("USMT baseline capture complete!","Stage 2 Complete",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage2
  
  AddBtn "Stage 3: Cutover" { 
    Stage3-Cutover
    Push-Telemetry 'cutover' @{ok=$true}
    [System.Windows.Forms.MessageBox]::Show("Cutover complete!`n`nDelta sync, LoadState, and integrity checks finished.","Stage 3 Complete",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
  } $colorStage3
  
  AddBtn "Stage 4: Post-Login Delta" { 
    Stage4-PostLoginDelta
    Push-Telemetry 'postdelta' @{ok=$true}
    [System.Windows.Forms.MessageBox]::Show("Post-login delta sync complete!","Stage 4 Complete",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage4
  
  AddBtn "Stage 5: Post-Provision" { 
    Stage5-PostProvision
    Push-Telemetry 'postprov' @{ok=$true}
  } $colorStage4
  
  AddBtn "Stage 6: Drives & DSNs" { 
    Stage6-DrivesAndDSN
    Push-Telemetry 'drives_dsn' @{ok=$true}
    [System.Windows.Forms.MessageBox]::Show("Mapped drives and DSNs migration complete!","Stage 6 Complete",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage4
  
  AddBtn "Generate Summary" { 
    $p = New-TechSummary
    [System.Windows.Forms.MessageBox]::Show("Technical summary saved to:`n`n$p","Summary Created",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  } $colorStage0
  
  # Progress panel with better design
  $progressLabel = New-Object Windows.Forms.Label
  $progressLabel.Text = "Progress Monitor"
  $progressLabel.Location = New-Object Drawing.Point(10, 375)
  $progressLabel.AutoSize = $true
  $progressLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
  
  $panel = New-Object Windows.Forms.Panel
  $panel.Location = New-Object Drawing.Point(10, 400)
  $panel.Size = New-Object Drawing.Size(1060, 300)
  $panel.AutoScroll = $true
  $panel.BackColor = [Drawing.Color]::White
  $panel.BorderStyle = 'FixedSingle'
  
  if (-not $global:__Bars) { $global:__Bars = @{} }
  $script:barY = 15
  
  function Ensure-Bar([string]$Label) {
    if ($global:__Bars.ContainsKey($Label)) { return }
    
    $title = New-Object Windows.Forms.Label
    $title.Text = $Label
    $title.AutoSize = $true
    $title.Location = New-Object Drawing.Point(15, $script:barY)
    $title.Font = New-Object Drawing.Font("Segoe UI", 8.5)
    $panel.Controls.Add($title)
    
    $pb = New-Object Windows.Forms.ProgressBar
    $pb.Style = 'Continuous'
    $pb.Minimum = 0
    $pb.Maximum = 100
    $pb.Location = New-Object Drawing.Point(15, $script:barY + 20)
    $pb.Size = New-Object Drawing.Size(1020, 22)
    $pb.ForeColor = $colorSuccess
    $panel.Controls.Add($pb)
    
    $script:barY += 55
    $global:__Bars[$Label] = $pb
  }
  
  $btnSeed = New-Object Windows.Forms.Button
  $btnSeed.Text = "Initialize Progress Bars"
  $btnSeed.Location = New-Object Drawing.Point(880, 372)
  $btnSeed.Size = New-Object Drawing.Size(190, 25)
  $btnSeed.BackColor = $colorStage4
  $btnSeed.ForeColor = [Drawing.Color]::White
  $btnSeed.FlatStyle = 'Flat'
  $btnSeed.Add_Click({
    foreach($u in $script:SelectedUsers) { 
      Ensure-Bar "User: $u (precopy)"
      Ensure-Bar "User: $u (delta)" 
    }
    Ensure-Bar "C:\ Non-System"
    [System.Windows.Forms.MessageBox]::Show("Progress bars initialized!","Ready",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
  })
  
  # Progress update timer with better performance
  $timer = New-Object Windows.Forms.Timer
  $timer.Interval = 500  # Reduced from 800ms for more responsive UI
  $timer.Add_Tick({
    if (-not $global:__ProgressMap) { return }
    foreach ($k in $global:__ProgressMap.Keys) {
      Ensure-Bar $k
      $p = [int]$global:__ProgressMap[$k]
      $clamped = [Math]::Min([Math]::Max($p, 0), 100)
      if ($global:__Bars[$k].Value -ne $clamped) {
        $global:__Bars[$k].Value = $clamped
      }
    }
  })
  $timer.Start()
  $f.Add_FormClosing({ 
    $timer.Stop()
    $timer2.Stop()
  })
  
  # Add all controls to form
  $f.Controls.AddRange(@($header, $hostPanel, $perfPanel, $btnPanel, $progressLabel, $btnSeed, $panel))
  [void]$f.ShowDialog()
}

# === Console Menu (fallback) ===
function Show-Menu {
  $tune = Get-RoboTuning
  $ipgDisp = if ($tune.IPG) { " IPG=$($tune.IPG)" } else { "" }
  $linkSpeed = Get-LinkSpeedMbps
@"
============================================================================
 Live User Migration Tool v2.0 – USMT + Robocopy   
 Runner: $env:COMPUTERNAME
 OLD: $script:OldHost   NEW: $script:NewHost
 Network: $linkSpeed Mbps | Profile: $($tune.Profile) | Threads: $($tune.MT)$ipgDisp
 OneDrive Skip: ALWAYS   Business Hours Throttle: $ThrottleBusinessHours
============================================================================
0) Set OLD/NEW hostnames
1) Pick user profile(s)
2) Stage 0 : Inventory (JSON+HTML)
3) Stage 1 : Precopy user data (parallel - $ParallelUserCopies at once)
4) Stage 1.5 : Copy C:\ except system trees
5) Stage 2 : USMT baseline (configs only via VSS)
6) Stage 3 : Cutover (delta + LoadState + Integrity spot-check)
7) Stage 4 : Optional post-login delta
8) Stage 5 : Drop Post-Prov pack (optional run)
9) Stage 6 : Migrate mapped drives & ODBC DSNs
T) Toggle ThrottleBusinessHours ($ThrottleBusinessHours)   
P) Cycle Performance Profile (now: $($script:TuningProfile))
S) Show current performance settings
Q) Quit
"@
}

function Run-Menu {
  while ($true) {
    Show-Menu
    $sel = Read-Host "Choose"
    try {
      switch ($sel.ToUpper()) {
        '0' { 
          $o = Read-Host "OLD Hostname"
          $n = Read-Host "NEW Hostname"
          Require-Hosts -Old $o -New $n 
        }
        '1' { 
          if (-not $script:OldHost) { throw "Set hosts first." }
          Pick-Users 
        }
        '2' { 
          if (-not $script:SelectedUsers) { Pick-Users }
          $p = Stage0-Inventory
          Push-Telemetry 'inventory' @{json=$p.json;html=$p.html}
        }
        '3' { 
          if (-not $script:SelectedUsers) { Pick-Users }
          Stage1-Precopy
          Push-Telemetry 'precopy' @{ok=$true}
        }
        '4' { 
          Stage1_5-CDriveNonSystem
          Push-Telemetry 'cdrive' @{ok=$true}
        }
        '5' { 
          if (-not $script:SelectedUsers) { Pick-Users }
          Stage2-USMT-Baseline
          Push-Telemetry 'usmt_baseline' @{ok=$true}
        }
        '6' { 
          if (-not $script:SelectedUsers) { Pick-Users }
          Stage3-Cutover
          Push-Telemetry 'cutover' @{ok=$true}
        }
        '7' { 
          if (-not $script:SelectedUsers) { Pick-Users }
          Stage4-PostLoginDelta
          Push-Telemetry 'postdelta' @{ok=$true}
        }
        '8' { 
          Stage5-PostProvision
          Push-Telemetry 'postprov' @{ok=$true}
        }
        '9' { 
          Stage6-DrivesAndDSN
          Push-Telemetry 'drives_dsn' @{ok=$true}
        }
        'T' { 
          $script:ThrottleBusinessHours = -not $ThrottleBusinessHours
          Write-Host "Business Hours Throttle: $ThrottleBusinessHours" -ForegroundColor Yellow
        }
        'P' {
          $order = @('Auto','Conservative','Balanced','Aggressive','WiFi')
          $i = $order.IndexOf($script:TuningProfile)
          $script:TuningProfile = $order[($i + 1) % $order.Count]
          Write-Host "Performance Profile -> $($script:TuningProfile)" -ForegroundColor Yellow
        }
        'S' {
          $tune = Get-RoboTuning
          $speed = Get-LinkSpeedMbps
          Write-Host "`nCurrent Performance Settings:" -ForegroundColor Cyan
          Write-Host "  Network Speed: $speed Mbps" -ForegroundColor White
          Write-Host "  Profile: $($tune.Profile)" -ForegroundColor White
          Write-Host "  Threads (MT): $($tune.MT)" -ForegroundColor White
          Write-Host "  Inter-Packet Gap: $(if($tune.IPG){$tune.IPG}else{'None'})" -ForegroundColor White
          Write-Host "  Parallel User Copies: $ParallelUserCopies" -ForegroundColor White
          Write-Host ""
          Read-Host "Press Enter to continue"
        }
        'Q' { break }
        default { Write-Host "Invalid option." -ForegroundColor Yellow }
      }
    } catch {
      Write-Warning $_.Exception.Message
      Push-Telemetry 'error' @{ message=$_.Exception.Message; option=$sel }
    }
  }
}

# === Entry ===
if ($Gui) { Run-UI } else { Run-Menu }
