# ============================================
# Endpoint Imaging Setup (domain creds + escrow UNC; UNC copy; robust cleanup)
# Steps:
# 1 mount, 2 copy, 3 run scripts (with admin member checker), 4 shortcut, 5 timezone, 6 asset reg,
# 7 cipher fix (always), 8 bitlocker (2-phase + final status), 9 gpupdate, summary + restart/exit
# ============================================

param(
  [switch]$RequireADEscrow   # require AD escrow success before enabling BitLocker
)

# --- Step 0: Sanity / Admin check ---
$currUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currUser)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Warning "Please run PowerShell as Administrator. Press any key to continue..."
  $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Paths
$UncSource   = "\\fileserver01\DeploymentShare\ImageFiles"     # <-- corrected UNC source
$EscrowShare = "\\fileserver01\DeploymentShare\KeyEscrow\"      # <-- hardcoded escrow UNC
$PsDriveSrc  = "SRC"
$PsDriveEsc  = "ESC"

$DestImage = "C:\Image Files"
$AddAdmin  = Join-Path $DestImage "Add-AdminGroups.ps1"
$ODvbs     = Join-Path $DestImage "OneDrive_Config_WIN10.vbs"
$Pilgrim   = Join-Path $DestImage "New Pilgrim.url"
$PubDesk   = [Environment]::GetFolderPath('CommonDesktopDirectory')
$PubPilgrim= Join-Path $PubDesk "New Pilgrim.url"

# Logs + status tracking
$LogsDir = Join-Path $env:ProgramData "Endpoint-Imaging-Logs"
New-Item -ItemType Directory -Force -Path $LogsDir | Out-Null
$RunId   = Get-Date -Format 'yyyyMMdd-HHmmss'
$MainLog = Join-Path $LogsDir "Run-$RunId.log"

$StepResults = [System.Collections.Generic.List[object]]::new()
function Write-Log([string]$msg, [string]$level = "INFO") {
  $line = "{0} [{1}] {2}" -f (Get-Date -Format s), $level, $msg
  $line | Tee-Object -FilePath $MainLog -Append | Out-Null
}
function Set-StepStatus([int]$num,[string]$name,[string]$status,[string]$detail="") {
  $StepResults.Add([pscustomobject]@{ Step=$num; Name=$name; Status=$status; Detail=$detail })
  Write-Log ("Step {0} - {1}: {2}{3}" -f $num,$name,$status,($(if($detail){" - $detail"}else{""})))
}
function Show-Step([int]$Step,[string]$Name) { Write-Host ("Step {0}: {1}" -f $Step, $Name) }
function Backup-RegistryKey {
  param([Parameter(Mandatory)] [string]$KeyPath,[Parameter(Mandatory)] [string]$OutReg)
  try { & reg.exe export $KeyPath $OutReg /y | Out-Null; Write-Log "Backed up '$KeyPath' to '$OutReg'" }
  catch { Write-Log "Failed to back up '$KeyPath': $_" "WARN" }
}

# --- Credential prompt (prefill CONTOSO\ but editable) ---
function Get-DomainCredential {
  param([string]$PrefillDomain = "CONTOSO")
  $prefilled = "{0}\" -f $PrefillDomain
  return Get-Credential -Message "Enter DOMAIN\username and password (default domain shown below; you can edit it)" -UserName $prefilled
}

# --- Helpers ---
function Invoke-GpUpdate {
  Write-Log "Running gpupdate /force ..."
  try {
    $p = Start-Process -FilePath gpupdate.exe -ArgumentList '/force' -Wait -PassThru -WindowStyle Hidden
    Write-Log "gpupdate exit code: $($p.ExitCode)"
    return $p.ExitCode
  } catch { Write-Log "gpupdate failed: $_" "WARN"; return 1 }
}

# Cipher Fix (REG_MULTI_SZ 'Functions' per your list)
function Apply-LegacySqlCipherFix {
  $sslBaseKey = 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL'
  $sslLeafKey = Join-Path $sslBaseKey '00010002'
  $backup = Join-Path $LogsDir "SSL-Backup-$RunId.reg"

  $cipherCsv = 'TLS_AES_256_GCM_SHA384,TLS_AES_128_GCM_SHA256,TLS_ECDHE_ECDSA_WITH_AES_256_GCM_SHA384,TLS_ECDHE_ECDSA_WITH_AES_128_GCM_SHA256,TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384,TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256,TLS_DHE_RSA_WITH_AES_256_GCM_SHA384,TLS_DHE_RSA_WITH_AES_128_GCM_SHA256,TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA384,TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA256,TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA384,TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA256,TLS_ECDHE_ECDSA_WITH_AES_256_CBC_SHA,TLS_ECDHE_ECDSA_WITH_AES_128_CBC_SHA,TLS_ECDHE_RSA_WITH_AES_256_CBC_SHA,TLS_ECDHE_RSA_WITH_AES_128_CBC_SHA,TLS_RSA_WITH_AES_256_GCM_SHA384,TLS_RSA_WITH_AES_128_GCM_SHA256,TLS_RSA_WITH_AES_256_CBC_SHA256,TLS_RSA_WITH_AES_128_CBC_SHA256,TLS_RSA_WITH_AES_256_CBC_SHA,TLS_RSA_WITH_AES_128_CBC_SHA,TLS_RSA_WITH_3DES_EDE_CBC_SHA,TLS_RSA_WITH_NULL_SHA256,TLS_RSA_WITH_NULL_SHA,TLS_PSK_WITH_AES_256_GCM_SHA384,TLS_PSK_WITH_AES_128_GCM_SHA256,TLS_PSK_WITH_AES_256_CBC_SHA384,TLS_PSK_WITH_AES_128_CBC_SHA256,TLS_PSK_WITH_NULL_SHA384,TLS_PSK_WITH_NULL_SHA256'
  $cipherArray = $cipherCsv.Split(',').ForEach({ $_.Trim() }) | Where-Object { $_ -ne '' }

  Backup-RegistryKey -KeyPath 'HKLM\SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL' -OutReg $backup

  if (-not (Test-Path -LiteralPath "Registry::$sslLeafKey")) {
    New-Item -Path "Registry::$sslLeafKey" -Force | Out-Null
    Write-Log "Created $sslLeafKey"
  }

  New-ItemProperty -Path "Registry::$sslLeafKey" -Name 'Functions' -Value $cipherArray -PropertyType MultiString -Force | Out-Null

  $post = (Get-ItemProperty -Path "Registry::$sslLeafKey" -Name 'Functions' -ErrorAction Stop).Functions
  return @($post).Count
}

# BitLocker (2-phase) with hardcoded escrow UNC and domain creds
function Ensure-BitLockerEnabled {
  param(
    [string]$MountPoint = 'C:',
    [string]$RecoveryDir = (Join-Path $env:ProgramData "Endpoint-Imaging-Logs"),
    [switch]$RequireADEscrow,
    [pscredential]$NetworkCredential
  )
  New-Item -ItemType Directory -Force -Path $RecoveryDir | Out-Null

  if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
    return @{ result="SKIPPED"; reason="BitLocker cmdlets unavailable" }
  }

  try { $vol = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop } catch {
    return @{ result="ERROR"; reason="Get-BitLockerVolume failed: $_" }
  }
  if ($vol.ProtectionStatus -eq 'On') { return @{ result="OK"; reason="Already enabled" } }

  try { $tpm = Get-Tpm } catch { $tpm = $null }
  if (-not $tpm -or -not $tpm.TpmPresent -or -not $tpm.TpmReady) {
    return @{ result="SKIPPED"; reason="TPM not present/ready" }
  }

  # Phase A: add Recovery Password protector
  try {
    $kp = Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector -ErrorAction Stop
  } catch {
    return @{ result="ERROR"; reason="Add-BitLockerKeyProtector failed: $_" }
  }
  $recPwd   = $kp.RecoveryPasswordProtector.RecoveryPassword
  $id       = $kp.KeyProtectorId
  $host     = $env:COMPUTERNAME
  $timestamp= (Get-Date).ToString('yyyyMMdd-HHmmss')

  # Save locally
  $localFile = Join-Path $RecoveryDir "BitLocker-$($MountPoint.TrimEnd(':'))-Recovery-$timestamp.txt"
  try {
    @(
      "ComputerName=$host"
      "MountPoint=$MountPoint"
      "KeyProtectorId=$id"
      "RecoveryPassword=$recPwd"
    ) | Out-File -FilePath $localFile -Encoding ASCII -Force
    Write-Log "Saved local recovery file: $localFile"
  } catch { Write-Log "Failed writing local recovery file: $_" "WARN" }

  # Map escrow UNC with provided domain creds, write <KeyProtectorId>-<HOST>.txt
  $uncFile = $null
  try {
    if (Get-PSDrive -Name $PsDriveEsc -ErrorAction SilentlyContinue) { Remove-PSDrive -Name $PsDriveEsc -ErrorAction SilentlyContinue }
    New-PSDrive -Name $PsDriveEsc -PSProvider FileSystem -Root $EscrowShare -Credential $NetworkCredential -ErrorAction Stop | Out-Null
    if (Test-Path "$PsDriveEsc`:") {
      $uncFile = Join-Path "$PsDriveEsc`:" ("{0}-{1}.txt" -f $id, $host)
      @(
        "ComputerName=$host"
        "MountPoint=$MountPoint"
        "KeyProtectorId=$id"
        "RecoveryPassword=$recPwd"
      ) | Out-File -FilePath $uncFile -Encoding ASCII -Force
      Write-Log "Escrowed recovery password to $uncFile"
    } else {
      Write-Log "Escrow PSDrive not accessible after mount." "WARN"
    }
  } catch {
    Write-Log "Failed to escrow to UNC: $_" "WARN"
  } finally {
    try { Remove-PSDrive -Name $PsDriveEsc -ErrorAction SilentlyContinue } catch {}
  }

  # Proactive AD escrow
  $adBacked = $false
  try {
    $null = & manage-bde -protectors -adbackup $MountPoint -id $id 2>&1
    if ($LASTEXITCODE -eq 0) { $adBacked = $true; Write-Log "AD escrow succeeded for $id" } else { Write-Log "AD escrow exit $LASTEXITCODE for $id" "WARN" }
  } catch { Write-Log "AD escrow threw: $_" "WARN" }

  if ($RequireADEscrow -and -not $adBacked) {
    return @{ result="SKIPPED"; reason="AD escrow not yet confirmed (RequireADEscrow set)" }
  }

  # Phase B: enable BitLocker
  try {
    Enable-BitLocker -MountPoint $MountPoint -UsedSpaceOnly -TpmProtector -EncryptionMethod XtsAes256 -SkipHardwareTest -ErrorAction Stop
  } catch {
    return @{ result="ERROR"; reason="Enable-BitLocker failed: $_" }
  }

  $note = "Enabled; Local=$localFile" + ($(if($uncFile){"; UNC=$uncFile"}else{"; UNC escrow failed"})) + ($(if($adBacked){"; AD escrow ok"}else{"; AD escrow pending"}))
  return @{ result="OK"; reason=$note }
}

# =========================
# Step 1: Auth & mount source
# =========================
$step = 1; $name = "Mount Image Source as PSDrive"
Show-Step $step $name
try {
  $cred = Get-DomainCredential
  if (Get-PSDrive -Name $PsDriveSrc -ErrorAction SilentlyContinue) { Remove-PSDrive -Name $PsDriveSrc -ErrorAction SilentlyContinue }
  New-PSDrive -Name $PsDriveSrc -PSProvider FileSystem -Root $UncSource -Credential $cred -ErrorAction Stop | Out-Null
  if (-not (Test-Path "$PsDriveSrc`:\")) { throw "Mount failed." }
  Set-StepStatus $step $name "OK" "Mounted $UncSource as $PsDriveSrc`:\"
} catch {
  Set-StepStatus $step $name "ERROR" "$_"
  Pause; exit 1
}
Pause

# =========================
# Step 2: Copy Image Files (use UNC directly)
# =========================
$step = 2; $name = "Copy to C:\Image Files"
Show-Step $step $name
try {
  if (-not (Test-Path -LiteralPath $DestImage)) {
    New-Item -Path $DestImage -ItemType Directory -Force | Out-Null
  }

  if (-not (Test-Path -LiteralPath $UncSource)) {
    throw "Source not reachable: $UncSource"
  }

  $log = Join-Path $env:TEMP "ImageFilesCopy-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

  & robocopy "$UncSource" "$DestImage" /E /R:2 /W:3 /NFL /NDL /NP "/LOG:$log"
  $rc = $LASTEXITCODE

  if ($rc -gt 7) {
    Set-StepStatus $step $name "WARN" "Robocopy exit $rc (Log: $log)"
  } else {
    Set-StepStatus $step $name "OK" "Log: $log"
  }
} catch {
  Set-StepStatus $step $name "ERROR" "$_"
}
Pause

# Ensure we stay on a local path the rest of the run (avoid UNC CWD issues later)
try { Set-Location -Path $env:SystemRoot } catch { Set-Location -Path 'C:\' }

# =========================
# Step 3: Run scripts (with Local Admins membership checker)
# =========================
$step = 3; $name = "Ensure Local Admin groups + OneDrive VBScript"
Show-Step $step $name
$detail = @()
try {
  # Targets to ensure in Local Administrators
  $targetAdmins = @(
    "CONTOSO\EndpointWorkstationAdmins",
    "CONTOSO\ServiceDeskOperators",
    "CONTOSO\FieldITSpecialists",
    "CONTOSO\OnsiteSupport",
    "CONTOSO\TemporaryAdmins"
  )

  # Current members (names come as DOMAIN\Name or MACHINE\Name)
  $current = @()
  try {
    $current = (Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop | Select-Object -ExpandProperty Name) | ForEach-Object { $_.ToLower() }
  } catch {
    Write-Log "Get-LocalGroupMember failed: $_" "WARN"
    $current = @()
  }

  $lowerTargets = $targetAdmins | ForEach-Object { $_.ToLower() }
  $missing = @($lowerTargets | Where-Object { $_ -notin $current })

  if ($missing.Count -eq 0) {
    $detail += "Admins OK (all present)"
    $skippedAddAdminScript = $true
  } else {
    $detail += ("Missing admins: {0}" -f ($missing -join ", "))
    $added = @()
    $failed = @()
    foreach ($m in $missing) {
      try {
        Add-LocalGroupMember -Group "Administrators" -Member $m -ErrorAction Stop
        $added += $m
      } catch {
        $failed += $m
        Write-Log ("Add-LocalGroupMember failed for {0}: {1}" -f $m, $_) "WARN"
      }
    }
    if ($failed.Count -gt 0 -and (Test-Path $AddAdmin)) {
      # Fallback: run external script
      & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $AddAdmin
      $detail += ("Fallback script ran; added={0}; failed={1}" -f ($added -join ","), ($failed -join ","))
    } else {
      if ($added.Count -gt 0) { $detail += ("Added inline: {0}" -f ($added -join ",")) }
      if ($failed.Count -gt 0) { $detail += ("Still failed: {0}" -f ($failed -join ",")) }
    }
  }

  # OneDrive script (independent of admin adds)
  if (Test-Path $ODvbs) { & cscript.exe //nologo $ODvbs; $detail += "OneDrive VBS OK" } else { $detail += "OneDrive VBS MISSING" }

  # Determine status for Step 3
  $status =
    if ($missing.Count -eq 0 -and $detail -notmatch "MISSING") { "OK" }
    elseif ($detail -match "Still failed") { "WARN" }
    elseif ($detail -match "MISSING") { "WARN" }
    else { "OK" }

  # Note if we skipped running the external Add Admin script
  if ($skippedAddAdminScript) { $detail += "Skipped Add-AdminGroups.ps1 (already present)" }

  Set-StepStatus $step $name $status ($detail -join "; ")

} catch {
  Set-StepStatus $step $name "ERROR" "$_"
}
Pause

# =========================
# Step 4: Copy shortcut
# =========================
$step = 4; $name = "Copy New Pilgrim.url to Public Desktop"
Show-Step $step $name
try {
  if (Test-Path $Pilgrim) { Copy-Item -LiteralPath $Pilgrim -Destination $PubPilgrim -Force; Set-StepStatus $step $name "OK" $PubPilgrim }
  else { Set-StepStatus $step $name "WARN" "Source not found: $Pilgrim" }
} catch { Set-StepStatus $step $name "ERROR" "$_" }
Pause

# =========================
# Step 5: Timezone + sync
# =========================
$step = 5; $name = "Set PST timezone + sync time"
Show-Step $step $name
try {
  tzutil /s "Pacific Standard Time" | Out-Null
  w32tm /resync /force | Out-Null
  Set-StepStatus $step $name "OK" "PST + w32tm resync"
} catch { Set-StepStatus $step $name "WARN" "$_" }
Pause

# =========================
# Step 6: Registry AssetNumber
# =========================
$step = 6; $name = "Write HKLM:\SOFTWARE\ESI\AssetNumber"
Show-Step $step $name
try {
  $asset = Read-Host "Enter Asset Number"
  $regPath = "HKLM:\SOFTWARE\ESI"
  if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
  New-ItemProperty -Path $regPath -Name "AssetNumber" -Value $asset -PropertyType String -Force | Out-Null
  Set-StepStatus $step $name "OK" "AssetNumber=$asset"
} catch { Set-StepStatus $step $name "ERROR" "$_" }
Pause

# =========================
# Step 7: Cipher Fix (always)
# =========================
$step = 7; $name = "Apply SQL/ODBC TLS cipher fix"
Show-Step $step $name
try {
  $count = Apply-LegacySqlCipherFix
  if ($count -ge 1) { Set-StepStatus $step $name "OK" "Functions entries: $count" }
  else { Set-StepStatus $step $name "WARN" "No Functions entries detected post-apply" }
} catch { Set-StepStatus $step $name "ERROR" "$_" }
Pause

# =========================
# Step 8: BitLocker 2-phase (uses domain creds for UNC escrow) + final status
# =========================
$step = 8; $name = "BitLocker check/enable on C:"
Show-Step $step $name
try {
  $bl = Ensure-BitLockerEnabled -MountPoint 'C:' -RecoveryDir $LogsDir -RequireADEscrow:$RequireADEscrow -NetworkCredential $cred

  # Always try to fetch final status, even if skipped or error
  $finalStatus = $null
  try {
    $vol = Get-BitLockerVolume -MountPoint 'C:' -ErrorAction Stop
    $finalStatus = "Protection=$($vol.ProtectionStatus); EncryptionPercent=$([int]$vol.EncryptionPercentage)%"
  } catch {
    $finalStatus = "Unable to query final BitLocker status: $_"
  }

  switch ($bl.result) {
    "OK"      { Set-StepStatus $step $name "OK"      ("{0}; {1}" -f $bl.reason,$finalStatus) }
    "SKIPPED" { Set-StepStatus $step $name "SKIPPED" ("{0}; {1}" -f $bl.reason,$finalStatus) }
    "ERROR"   { Set-StepStatus $step $name "ERROR"   ("{0}; {1}" -f $bl.reason,$finalStatus) }
    default   { Set-StepStatus $step $name "WARN"    ("Unknown result; {0}" -f $finalStatus) }
  }
} catch {
  Set-StepStatus $step $name "ERROR" "$_"
}
Pause

# =========================
# Step 9: gpupdate /force
# =========================
$step = 9; $name = "Group Policy refresh"
Show-Step $step $name
try {
  $code = Invoke-GpUpdate
  if ($code -eq 0) { Set-StepStatus $step $name "OK" "gpupdate exit 0" }
  else { Set-StepStatus $step $name "WARN" "gpupdate exit $code" }
} catch { Set-StepStatus $step $name "ERROR" "$_" }
Pause

# --- Cleanup source/escrow mappings (robust) ---
Write-Host "Cleaning up mapped drives..."
try {
  # Ensure we're on a local path (avoid UNC CWD issues with cmd.exe)
  Set-Location -Path $env:SystemRoot
} catch {
  Set-Location -Path 'C:\'
}

# Remove our PSDrives if present
try { Remove-PSDrive -Name $PsDriveSrc -ErrorAction SilentlyContinue } catch {}
try { Remove-PSDrive -Name $PsDriveEsc -ErrorAction SilentlyContinue } catch {}

# Clear SMB mappings to \\fileserver01 if any remain
$server = '\\fileserver01'
$cleared = $false

if (Get-Command -Name Get-SmbMapping -ErrorAction SilentlyContinue) {
  try {
    $maps = Get-SmbMapping | Where-Object { $_.RemotePath -like "$server*" }
    foreach ($m in $maps) {
      try {
        Remove-SmbMapping -RemotePath $m.RemotePath -Force -UpdateProfile -ErrorAction Stop
        Write-Log "Removed SMB mapping: $($m.RemotePath)"
        $cleared = $true
      } catch {
        Write-Log "Failed Remove-SmbMapping $($m.RemotePath): $_" "WARN"
      }
    }
  } catch {
    Write-Log "Get-SmbMapping failed: $_" "WARN"
  }
}

# Fallback to NET USE only if a session to the server exists
try {
  $netOut = cmd.exe /c "net use"
  if ($netOut -match [regex]::Escape($server)) {
    $null = cmd.exe /c "net use $server /delete /y"
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 2250) {
      Write-Log "NET USE session to $server cleared (exit $LASTEXITCODE)"
      $cleared = $true
    } else {
      Write-Log "NET USE delete returned $LASTEXITCODE" "WARN"
    }
  } else {
    Write-Log "No NET USE session for $server"
  }
} catch {
  Write-Log "NET USE cleanup threw: $_" "WARN"
}

Write-Log ("Cleanup complete." + ($(if($cleared){" (sessions cleared)"}else{" (no sessions found)"})))
Write-Host "Cleanup complete."

# =========================
# Summary + Restart/Exit
# =========================
Write-Host ""
Write-Host "========== Endpoint Imaging Summary =========="
$StepResults | Sort-Object Step | ForEach-Object {
  "{0}. {1,-35} : {2}{3}" -f $_.Step,$_.Name,$_.Status,($(if ($_.Detail) { " - $($_.Detail)" } else { "" }))
}
Write-Host "========================================="
Write-Host ("Full log: {0}" -f $MainLog)
Write-Host ""

# Offer Restart or Exit
$choices = New-Object System.Collections.ObjectModel.Collection[System.Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object System.Management.Automation.Host.ChoiceDescription "&Restart","Restart the computer now."))
$choices.Add((New-Object System.Management.Automation.Host.ChoiceDescription "E&xit","Exit without restart."))
$selection = $Host.UI.PromptForChoice("Actions","Choose what to do next:",$choices,1)
if ($selection -eq 0) {
  Write-Host "Restarting now..."
  Restart-Computer -Force
} else {
  Write-Host "Exiting without restart."
}
