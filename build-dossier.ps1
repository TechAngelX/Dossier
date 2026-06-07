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

# 3. Publish
Write-Host "[3/4] Building self-contained app..." -ForegroundColor Blue
dotnet publish -c Release -r $Runtime `
    --self-contained true `
    --nologo `
    -v q

$PublishDir = Join-Path $ScriptDir "bin\Release\net10.0\$Runtime\publish"
if (-not (Test-Path "$PublishDir\$OutputExe")) {
    $PublishDir = Join-Path $ScriptDir "bin\Release\$Runtime\publish"
}
if (-not (Test-Path "$PublishDir\$OutputExe")) {
    Write-Host "ERROR: could not find built executable at $PublishDir" -ForegroundColor Red
    exit 1
}

# 4. Copy to Desktop
Write-Host "[4/4] Copying to Desktop..." -ForegroundColor Blue
$Dest = Join-Path $Desktop $OutputExe
Copy-Item "$PublishDir\$OutputExe" $Dest -Force

$Size = [math]::Round((Get-Item $Dest).Length / 1MB, 1)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  Build Complete!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Location : $Dest" -ForegroundColor Blue
Write-Host "  Size     : ${Size} MB" -ForegroundColor Blue
Write-Host ""
Write-Host "  First-time setup on a new machine:" -ForegroundColor Yellow
Write-Host "    1. Run Dossier.exe once (it will fail on browser launch)" -ForegroundColor Yellow
Write-Host "    2. Open a terminal and run:  playwright install msedge" -ForegroundColor Yellow
Write-Host ""
Write-Host "Done!" -ForegroundColor Green
