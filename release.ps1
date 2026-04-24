# Release Workflow Script
# Wrapper around build.ps1 -> Uploads to GitHub

param(
    [Parameter(Mandatory = $true)]
    [string]$Version,
    [ValidateSet("Standard", "Small", "All")]
    [string]$Type = "All"
)

$ErrorActionPreference = "Stop"

# 1. Build via centralized script
Write-Host "Starting Release v$Version..." -ForegroundColor Cyan
.\build.ps1 -Version $Version -Type $Type -OutputBaseDir "bin\Publish"

if ($LASTEXITCODE -ne 0) { exit 1 }

# 2. Check Authentication
gh auth status 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Please authenticate with GitHub CLI first: gh auth login" -ForegroundColor Red
    exit 1
}

# 3. Clean 'latest' release
Write-Host "`nDeleting existing 'latest' release..." -ForegroundColor Yellow
gh release delete latest --yes --cleanup-tag 2>$null

# 4. Upload Logic
Write-Host "Creating release..." -ForegroundColor Cyan

# Collect assets
$assets = @()

if ($Type -eq "Standard" -or $Type -eq "All") {
    $path = "bin\Publish\Standard\ExcelPowerPivotMcp.exe"
    if (Test-Path $path) {
        $assets += $path
    }
}

if ($Type -eq "Small" -or $Type -eq "All") {
    $path = "bin\Publish\Small\ExcelPowerPivotMcp.exe"
    if (Test-Path $path) {
        # Rename for upload clarity if needed, or upload as-is but with a distinct label logic?
        # GitHub releases need unique filenames.
        # We must rename one if they have the same name!
        $uploadPath = "bin\Publish\Small\ExcelPowerPivotMcp-Small.exe"
        Copy-Item $path $uploadPath -Force
        $assets += $uploadPath
    }
}

$notes = @"
## Excel Power Pivot MCP Server v$Version

### Downloads
- **ExcelPowerPivotMcp.exe** (~67MB)
  - *Standard Version*. Portable, includes everything needed.
  
- **ExcelPowerPivotMcp-Small.exe** (~5MB)
  - *Small Version*. Requires [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0).

See [README](https://github.com/back1ply/Excel-Power-Pivot-MCP#installation) for setup instructions.
"@

# Create and upload
gh release create latest `
    --title "v$Version" `
    --notes $notes `
    --latest `
    $assets

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nRelease v$Version published successfully!" -ForegroundColor Green
    Write-Host "View at: https://github.com/back1ply/Excel-Power-Pivot-MCP/releases/latest" -ForegroundColor Cyan
}
else {
    Write-Host "Release upload failed!" -ForegroundColor Red
    exit 1
}
