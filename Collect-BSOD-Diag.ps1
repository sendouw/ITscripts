# Collect-BSOD-Diag.ps1
#run with this command:
#Set-ExecutionPolicy Bypass -Scope Process -Force .\Collect-BSOD-Diag.ps1
# Saves BSOD diagnostic data to C:\Temp\BSOD_Diag_<timestamp>

$ts = Get-Date -Format "yyyy-MM-dd_HHmmss"
$outFolder = "C:\Temp\BSOD_Diag_$ts"
New-Item -ItemType Directory -Path $outFolder -Force | Out-Null

# 1. System Info
systeminfo > "$outFolder\SystemInfo.txt"

# 2. BIOS Version
Get-CimInstance -ClassName Win32_BIOS | Out-File "$outFolder\BIOS.txt"

# 3. Installed Drivers
Get-WmiObject Win32_PnPSignedDriver |
  Sort-Object DriverDate -Descending |
  Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer, InfName |
  Out-File "$outFolder\Drivers.txt"

# 4. Filter Drivers (fltmgr)
fltmc filters > "$outFolder\FltmcFilters.txt"

# 5. Running Services
Get-Service | Sort-Object Status | Out-File "$outFolder\Services.txt"

# 6. Installed Software
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
  Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
  Where-Object { $_.DisplayName } |
  Sort-Object DisplayName |
  Out-File "$outFolder\InstalledPrograms.txt"

# 7. Core Isolation / HVCI Status
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
$hvciStatus = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue | Select-Object Enabled
$hvciStatus | Out-File "$outFolder\HVCI.txt"

# 8. Output location
Write-Host "✅ Diagnostic data collected at: $outFolder" -ForegroundColor Green
