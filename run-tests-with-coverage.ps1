# Run tests with code coverage reporting
# Generates multiple output formats: opencover (XML), lcov, cobertura, and JSON
# HTML report can be generated using ReportGenerator tool

Write-Host "Running tests with code coverage..." -ForegroundColor Cyan

# Clean previous coverage results
if (Test-Path ".\TestResults") {
    Write-Host "Cleaning previous test results..." -ForegroundColor Yellow
    Remove-Item -Recurse -Force ".\TestResults"
}

# Run tests with coverage collection
dotnet test `
    /p:CollectCoverage=true `
    /p:CoverletOutputFormat="opencover%2clcov%2ccobertura%2cjson" `
    /p:CoverletOutput=".\TestResults\coverage" `
    /p:ExcludeByFile="**/*.g.cs%2c**/obj/**/*.cs" `
    /p:Exclude="[xunit.*]*%2c[*.Tests]*" `
    --logger "console;verbosity=normal"

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nTests passed! Coverage reports generated in .\TestResults\" -ForegroundColor Green
    Write-Host "  - coverage.opencover.xml (for ReportGenerator)" -ForegroundColor Gray
    Write-Host "  - coverage.info (LCOV format)" -ForegroundColor Gray
    Write-Host "  - coverage.cobertura.xml (Cobertura format)" -ForegroundColor Gray
    Write-Host "  - coverage.json (JSON format)" -ForegroundColor Gray

    # Check if ReportGenerator is installed
    $reportGen = Get-Command reportgenerator -ErrorAction SilentlyContinue
    if ($reportGen) {
        Write-Host "`nGenerating HTML coverage report..." -ForegroundColor Cyan
        reportgenerator `
            -reports:".\TestResults\coverage.opencover.xml" `
            -targetdir:".\TestResults\CoverageReport" `
            -reporttypes:"Html;Badges"

        Write-Host "HTML report generated at .\TestResults\CoverageReport\index.html" -ForegroundColor Green

        # Open the report in default browser
        $openReport = Read-Host "Open HTML report in browser? (Y/N)"
        if ($openReport -eq "Y" -or $openReport -eq "y") {
            Start-Process ".\TestResults\CoverageReport\index.html"
        }
    } else {
        Write-Host "`nTo generate HTML reports, install ReportGenerator:" -ForegroundColor Yellow
        Write-Host "  dotnet tool install --global dotnet-reportgenerator-globaltool" -ForegroundColor Gray
    }
} else {
    Write-Host "`nTests failed!" -ForegroundColor Red
    exit 1
}
