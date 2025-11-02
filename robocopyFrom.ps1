<# 
Menu-free, reliable source/destination selection:
- Prefer GUI FolderBrowserDialog for Source/Destination
- Fallback to Read-Host if GUI not available
- Robocopy /MT:64 with progress + logging
#>

$ErrorActionPreference = 'Stop'

function Pick-Folder([string]$title) {
    # Try GUI picker first
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = $title
        $fb.ShowNewFolderButton = $true
        $result = $fb.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $fb.SelectedPath) {
            return $fb.SelectedPath
        }
    } catch {
        # ignore, fallback below
    }

    # Fallback: Read-Host with validation
    while ($true) {
        $p = Read-Host "$title (type/paste full path)"
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $p = $p.Trim('"').Trim()
        if (Test-Path -LiteralPath $p) { return (Resolve-Path -LiteralPath $p).Path }
        Write-Host "Path not found: $p" -ForegroundColor Yellow
    }
}

function Explain-ExitCode([int]$code) {
    if ($code -ge 8) { return "Errors occurred (exit code $code). Review the log." }
    switch ($code) {
        0 { "No files were copied. Everything already up to date." }
        1 { "Files were copied successfully." }
        2 { "Extra files/dirs detected (check log)." }
        3 { "Files copied + extra detected (check log)." }
        5 { "Mismatched files; files copied (verify results)." }
        6 { "Extra + mismatched detected (check log)." }
        7 { "Files copied + extra + mismatched (check log)." }
        default { "Completed with exit code $code (see log)." }
    }
}

Clear-Host
Write-Host "Robocopy 64-thread helper (GUI picker)" -ForegroundColor Green
Write-Host "--------------------------------------"

$Source      = Pick-Folder "Select SOURCE folder"
$Destination = Pick-Folder "Select DESTINATION folder"

Write-Host ""
Write-Host "Source     : $Source"
Write-Host "Destination: $Destination"
Write-Host ""

$confirm = Read-Host "Proceed with copy? (Y/n)"
if ($confirm -match '^(n|no)$') { Write-Host "Aborted." -ForegroundColor Yellow; exit 0 }

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logPath   = Join-Path $env:USERPROFILE "Desktop\robocopy-$timestamp.log"
Write-Host "Logging to: $logPath" -ForegroundColor DarkCyan
Write-Host ""

# Build robocopy args
$arguments = @(
    $Source
    $Destination
    '/E'            # subdirs incl empty
    '/Z'            # restartable
    '/MT:64'        # 64 threads
    '/R:2','/W:2'   # quick retries
    '/COPY:DATSO'   # data+attrs+timestamps+ACLs+owner
    '/DCOPY:DAT'    # dirs with data/attrs/timestamps
    '/ETA'          # progress
    '/TEE'          # to console + log
    "/LOG:`"$logPath`""
    '/XD','System Volume Information','$RECYCLE.BIN'
)

Write-Host "Starting Robocopy..." -ForegroundColor Green
Write-Host "-------------------------------------------------------------"
& robocopy @arguments
$rc = $LASTEXITCODE

Write-Host "-------------------------------------------------------------"
Write-Host (Explain-ExitCode -code $rc) -ForegroundColor Cyan
Write-Host "Exit code: $rc"
Write-Host "Log saved to: $logPath"
