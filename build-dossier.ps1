# ============================================================
# Dossier - Build to Desktop Script (Windows PowerShell)
# ============================================================
# Run from the project root:  .\build-dossier.ps1
# Requires: .NET 10 SDK  (winget install Microsoft.DotNet.SDK.10)
#           Microsoft Edge  (already on Windows 11)
# After first run also run:  playwright install msedge
# ============================================================

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Dossier  --  Windows Build Script" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

$Runtime   = "win-x64"
$Desktop   = [Environment]::GetFolderPath("Desktop")
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$OutputExe = "Dossier.exe"

Write-Host "  Platform : Windows x64" -ForegroundColor Green
Write-Host "  Output   : $Desktop\$OutputExe" -ForegroundColor Green
Write-Host ""

Set-Location $ScriptDir

# Keep NuGet cache on a short local path to avoid Windows 260-char path limit
$env:NUGET_PACKAGES = "C:\np"

# 1. Clean
Write-Host "[1/4] Cleaning previous builds..." -ForegroundColor Blue
dotnet clean -c Release --nologo -v q 2>$null

# 2. Restore
Write-Host "[2/4] Restoring dependencies..." -ForegroundColor Blue
dotnet restore --nologo -v q

# 3. Publish (TRUE single file, self-contained)
# Everything - the .NET runtime AND Playwright's native driver (.playwright/) -
# is packed INSIDE Dossier.exe. IncludeAllContentForSelfExtract is the critical
# flag: it bundles the driver's content files (node.exe + driver package) into
# the exe. At launch the runtime transparently unpacks them to a hidden temp
# folder, where AppContext.BaseDirectory points, so Playwright finds its driver.
# (IncludeNativeLibrariesForSelfExtract alone does NOT include those files -
# that was the old bug: "Driver not found: C:\Users\.playwright\...".)
Write-Host "[3/4] Building self-contained single-file app..." -ForegroundColor Blue
# Publish the app project explicitly (not the .sln) - otherwise the test project
# gets dragged into single-file publish and fails with NETSDK1098.
dotnet publish Dossier.csproj -c Release -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeAllContentForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    --nologo `
    -v q
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: publish failed (exit $LASTEXITCODE)." -ForegroundColor Red
    exit 1
}

$PublishDir = Join-Path $ScriptDir "bin\Release\net10.0\$Runtime\publish"
if (-not (Test-Path "$PublishDir\$OutputExe")) {
    $PublishDir = Join-Path $ScriptDir "bin\Release\$Runtime\publish"
}
if (-not (Test-Path "$PublishDir\$OutputExe")) {
    Write-Host "ERROR: could not find built executable at $PublishDir" -ForegroundColor Red
    exit 1
}

# 4. Copy the single exe to the Desktop
# One file, everything inside. Also remove any old leftover Dossier folder from
# a previous directory-publish build so the Desktop stays clean.
Write-Host "[4/4] Copying to Desktop..." -ForegroundColor Blue
$Dest = Join-Path $Desktop $OutputExe

$StaleDir = Join-Path $Desktop "Dossier"
if (Test-Path $StaleDir) { Remove-Item $StaleDir -Recurse -Force }

Copy-Item "$PublishDir\$OutputExe" $Dest -Force

$Size = [math]::Round((Get-Item $Dest).Length / 1MB, 1)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  Build Complete!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Run this : $Dest" -ForegroundColor Blue
Write-Host "  Size     : ${Size} MB (one file)" -ForegroundColor Blue
Write-Host ""
Write-Host "  Requires Microsoft Edge (built into Windows 11 - nothing to install)." -ForegroundColor Yellow
Write-Host "  First launch unpacks bundled files to a temp folder and may take a" -ForegroundColor Yellow
Write-Host "  few extra seconds; later launches are fast." -ForegroundColor Yellow
Write-Host ""
Write-Host "Done!" -ForegroundColor Green