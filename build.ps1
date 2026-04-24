# Central Build Script
# Handles building both "Standard" (Self-Contained) and "Small" (Framework-Dependent) versions

param(
    [string]$Version = "1.0.0",
    [ValidateSet("Standard", "Small", "All")]
    [string]$Type = "Standard",
    [string]$OutputBaseDir = "bin\Publish"
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Excel Power Pivot MCP - Build v$Version" -ForegroundColor Cyan
Write-Host "  Type: $Type" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Clean
Write-Host "`n[1/3] Cleaning output directory ($OutputBaseDir)..." -ForegroundColor Yellow
if (Test-Path $OutputBaseDir) {
    Remove-Item -Path $OutputBaseDir -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -ItemType Directory -Path $OutputBaseDir -Force | Out-Null

# Function to build
function Build-Variant {
    param(
        [string]$VariantName,
        [bool]$SelfContained,
        [bool]$EnableCompression
    )
    
    $outDir = Join-Path $OutputBaseDir $VariantName
    Write-Host "`n  Building '$VariantName' variant..." -ForegroundColor Cyan
    Write-Host "    - SelfContained: $SelfContained" -ForegroundColor Gray
    Write-Host "    - Compression: $EnableCompression" -ForegroundColor Gray
    Write-Host "    - Output: $outDir" -ForegroundColor Gray

    $pCompression = if ($EnableCompression) { "/p:EnableCompressionInSingleFile=true" } else { "/p:EnableCompressionInSingleFile=false" }

    dotnet publish "ExcelPowerPivotMcp.csproj" -c Release -r win-x64 `
        --self-contained $SelfContained `
        -p:PublishSingleFile=true `
        $pCompression `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:Version=$Version `
        -p:DebugType=None -p:DebugSymbols=false `
        -o $outDir `
        --nologo


    if ($LASTEXITCODE -ne 0) {
        Write-Host "    FAILED!" -ForegroundColor Red
        exit 1
    }
    
    # Rename EXE for clarity in the output folder (optional, but helps if copied together)
    $exeName = "ExcelPowerPivotMcp.exe"
    if ($VariantName -eq "Small") {
        # Keep original name on disk but note it. 
        # Actually, let's just properly verify it exists.
    }
    
    $finalPath = Join-Path $outDir $exeName
    $sizeMb = "{0:N2}" -f ((Get-Item $finalPath).Length / 1MB)
    Write-Host "    SUCCESS. Size: $sizeMb MB" -ForegroundColor Green
}

# Build Process
Write-Host "`n[2/3] Building..." -ForegroundColor Yellow

if ($Type -eq "Standard" -or $Type -eq "All") {
    # Standard: Self-Contained, Compressed
    Build-Variant -VariantName "Standard" -SelfContained $true -EnableCompression $true
}

if ($Type -eq "Small" -or $Type -eq "All") {
    # Small: Framework-Dependent, No Compression (required for single-file)
    Build-Variant -VariantName "Small" -SelfContained $false -EnableCompression $false
}

# Summary
Write-Host "`n[3/3] Build Complete!" -ForegroundColor Green
Write-Host "Outputs in: $OutputBaseDir" -ForegroundColor White
Get-ChildItem -Path $OutputBaseDir -Recurse -Filter "*.exe" | ForEach-Object {
    $size = "{0:N2} MB" -f ($_.Length / 1MB)
    Write-Host "  - $($_.FullName) ($size)" -ForegroundColor Gray
}
