<#
.SYNOPSIS
    Builds NPTEL-Manager.exe from NPTEL-Manager.ps1 using ps2exe and packages
    a standalone Portable Zip Folder (NPTEL-Studio.zip) for coordinators.

.DESCRIPTION
    1. Compiles NPTEL-Manager.ps1 into NPTEL-Manager.exe (noConsole, STA, DPIAware, supportOS).
    2. Assembles clean distribution folder: dist\NPTEL-Studio\
    3. Bundles offline ImportExcel dependency into dist\NPTEL-Studio\modules\ImportExcel\
    4. Compresses into dist\NPTEL-Studio.zip (~15 MB) for 1-click coordinator execution.
#>

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " NPTEL Operations Studio -- Portable Coordinator Packaging  " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# 1. Ensure ps2exe is installed
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Installing ps2exe module..." -ForegroundColor Yellow
    Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
}
Import-Module ps2exe

# 2. Compile NPTEL-Manager.exe
$srcScript = Join-Path $root 'NPTEL-Manager.ps1'
$outExe    = Join-Path $root 'NPTEL-Manager.exe'

Write-Host "`n[1/4] Compiling NPTEL-Manager.exe..." -ForegroundColor Yellow

$ps2exeArgs = @{
    inputFile    = $srcScript
    outputFile   = $outExe
    noConsole    = $true
    STA          = $true
    DPIAware     = $true
    supportOS    = $true
    title        = 'NPTEL Management System'
    description  = 'Exception-based NPTEL Verification and Reporting Suite'
    product      = 'NPTEL Operations Studio'
    version      = '1.0.0.0'
    requireAdmin = $false
}

Invoke-ps2exe @ps2exeArgs

if (-not (Test-Path -LiteralPath $outExe)) {
    Write-Host "`nBuild failed: $outExe was not created." -ForegroundColor Red
    exit 1
}

$exeSizeMB = [Math]::Round(((Get-Item $outExe).Length / 1MB), 2)
Write-Host ("Compilation complete: {0} ({1} MB)" -f $outExe, $exeSizeMB) -ForegroundColor Green

# 3. Assemble Portable Distribution Folder
Write-Host ""
Write-Host "[2/4] Assembling portable distribution folder: dist\NPTEL-Studio\..." -ForegroundColor Yellow

$distRoot = Join-Path $root "dist"
$studioDir = Join-Path $distRoot "NPTEL-Studio"

if (Test-Path -LiteralPath $studioDir) {
    Remove-Item -LiteralPath $studioDir -Recurse -Force
}
$null = New-Item -ItemType Directory -Path $studioDir -Force

# 3.1 Copy Executable
Copy-Item -LiteralPath $outExe -Destination $studioDir -Force

# 3.2 Copy UI directory
Copy-Item -LiteralPath (Join-Path $root "UI") -Destination $studioDir -Recurse -Force

# 3.3 Copy modules directory (excluding any temporary test files)
$distModules = Join-Path $studioDir "modules"
$null = New-Item -ItemType Directory -Path $distModules -Force
Get-ChildItem -LiteralPath (Join-Path $root "modules") -Filter "*.ps1" | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $distModules -Force
}

# 3.4 Copy assets directory if present
$assetsDir = Join-Path $root "assets"
if (Test-Path -LiteralPath $assetsDir) {
    Copy-Item -LiteralPath $assetsDir -Destination $studioDir -Recurse -Force
}

# 3.5 Bundle ImportExcel module for 100% offline portability
Write-Host "`n[3/4] Bundling offline ImportExcel module..." -ForegroundColor Yellow
$sysImportExcel = (Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1).ModuleBase
if ($sysImportExcel -and (Test-Path -LiteralPath $sysImportExcel)) {
    $distImportExcel = Join-Path $distModules "ImportExcel"
    Copy-Item -LiteralPath $sysImportExcel -Destination $distImportExcel -Recurse -Force
    Write-Host "Bundled ImportExcel from $sysImportExcel" -ForegroundColor Green
} else {
    Write-Warning "ImportExcel module not found in system module paths. Coordinators will need internet access on first run."
}

# 3.6 Create clean starter data directory (zero mock data, clean empty state for coordinators)
$distData = Join-Path $studioDir "data"
$null = New-Item -ItemType Directory -Path $distData -Force
$distCoursesDir = Join-Path $distData "Courses"
$null = New-Item -ItemType Directory -Path $distCoursesDir -Force
Set-Content -LiteralPath (Join-Path $distData "courses.json") -Value "[]" -Encoding UTF8
$srcFormLinks = Join-Path $root "data\form_links.json"
if (Test-Path -LiteralPath $srcFormLinks) {
    Copy-Item -LiteralPath $srcFormLinks -Destination $distData -Force
}

# 3.7 Add Coordinator Quick Start Guide (README.txt)
$readmeContent = @"
========================================================================
 NPTEL Operations Studio -- Portable Coordinator Verification Suite
========================================================================

HOW TO RUN:
1. Extract the entire "NPTEL-Studio" folder to any convenient location on your PC (e.g. Desktop).
2. Double-click "NPTEL-Manager.exe" to launch the application.

REQUIREMENTS:
- Windows 10 or Windows 11 (64-bit).
- No installation, administrative privileges, or Microsoft Office required.
- Works 100% offline (All modules, UI views, and ImportExcel dependencies are pre-bundled).

INCLUDED FEATURES:
- Stage 1: Registration Verification & Fee Receipt Ingestion
- Stage 1 Verified Students Workspace (Search, 25/page Pagination, Receipt Preview, Unapprove/Undo staging, CSV export)
- Stage 2: Post-Exam Scorecard Ingestion & Certificate Mark Parsing (Assignments /25, Exam /75, Total /100)
- Stage 2 Verified Workspace & Dynamic Header Configurator (+ Add Header Info)
- Cumulative 4-Subsheet College Result Sheet Generator (Main Result, CW, End Sem, NPTEL)
- Offline Windows OCR & PDF rasterization (zero cloud API dependencies)
- Email Center for bulk discrepancy notifications with BCC student privacy
========================================================================
"@
Set-Content -LiteralPath (Join-Path $studioDir "README.txt") -Value $readmeContent -Encoding UTF8

# 4. Create Standalone Portable Zip Archive
Write-Host "`n[4/4] Compressing into dist\NPTEL-Studio.zip..." -ForegroundColor Yellow
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipPath = Join-Path $distRoot "NPTEL-Studio.zip"

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
[System.IO.Compression.ZipFile]::CreateFromDirectory($studioDir, $zipPath)

$zipSizeMB = [Math]::Round(((Get-Item $zipPath).Length / 1MB), 2)

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " BUILD AND PACKAGING SUCCESSFUL!                            " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Executable:       $studioDir\NPTEL-Manager.exe"
Write-Host "Portable Folder:  $studioDir"
Write-Host ("Shippable Zip:    {0} ({1} MB)" -f $zipPath, $zipSizeMB)
Write-Host "Coordinator can simply extract the zip and double-click NPTEL-Manager.exe!" -ForegroundColor Cyan
Write-Host "============================================================"
Write-Host ""
