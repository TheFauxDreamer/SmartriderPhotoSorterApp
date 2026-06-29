# ============================================================
#  Build the School Photo Organiser .exe
#  Run this ONCE on a Windows PC (that has internet access).
#  It creates "Photo Organiser.exe" next to this script.
#
#  HOW TO RUN:
#    Right-click this file > "Run with PowerShell"
#    (or open PowerShell, cd to this folder, and run  .\Build-PhotoOrganiser-EXE.ps1)
# ============================================================

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$inputFile  = Join-Path $scriptDir "Organise-Photos-GUI.ps1"
$outputFile = Join-Path $scriptDir "Photo Organiser.exe"
$iconFile   = Join-Path $scriptDir "icon.ico"   # optional - used only if present

Write-Host "Checking for the PS2EXE module..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Installing PS2EXE (one-time)..." -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

if (-not (Test-Path $inputFile)) {
    Write-Host "ERROR: Could not find Organise-Photos-GUI.ps1 next to this script." -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
}

Write-Host "Building '$outputFile'..." -ForegroundColor Cyan

# -noConsole = no black window;  -STA = required for the window to work
$params = @{
    InputFile  = $inputFile
    OutputFile = $outputFile
    noConsole  = $true
    STA        = $true
    title      = "School Photo Organiser"
    product    = "School Photo Organiser"
    description= "Sorts school photos into Staff / Left and reports missing photos"
}
if (Test-Path $iconFile) { $params.iconFile = $iconFile }

Invoke-ps2exe @params

Write-Host ""
if (Test-Path $outputFile) {
    Write-Host "Done! Created: $outputFile" -ForegroundColor Green
    Write-Host "You can now copy 'Photo Organiser.exe' to the photos folder and double-click it." -ForegroundColor Green
} else {
    Write-Host "Build did not produce an .exe - please check the messages above." -ForegroundColor Red
}
Write-Host ""
Read-Host "Press Enter to close"
