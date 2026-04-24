#!/bin/bash
# Run tests with code coverage reporting
# Generates multiple output formats: opencover (XML), lcov, cobertura, and JSON

echo "Running tests with code coverage..."

# Clean previous coverage results
if [ -d "./TestResults" ]; then
    echo "Cleaning previous test results..."
    rm -rf ./TestResults
fi

# Run tests with coverage collection
dotnet test \
    /p:CollectCoverage=true \
    /p:CoverletOutputFormat="opencover%2clcov%2ccobertura%2cjson" \
    /p:CoverletOutput="./TestResults/coverage" \
    /p:ExcludeByFile="**/*.g.cs%2c**/obj/**/*.cs" \
    /p:Exclude="[xunit.*]*%2c[*.Tests]*" \
    --logger "console;verbosity=normal"

if [ $? -eq 0 ]; then
    echo ""
    echo "Tests passed! Coverage reports generated in ./TestResults/"
    echo "  - coverage.opencover.xml (for ReportGenerator)"
    echo "  - coverage.info (LCOV format)"
    echo "  - coverage.cobertura.xml (Cobertura format)"
    echo "  - coverage.json (JSON format)"

    # Check if ReportGenerator is installed
    if command -v reportgenerator &> /dev/null; then
        echo ""
        echo "Generating HTML coverage report..."
        reportgenerator \
            -reports:"./TestResults/coverage.opencover.xml" \
            -targetdir:"./TestResults/CoverageReport" \
            -reporttypes:"Html;Badges"

        echo "HTML report generated at ./TestResults/CoverageReport/index.html"
    else
        echo ""
        echo "To generate HTML reports, install ReportGenerator:"
        echo "  dotnet tool install --global dotnet-reportgenerator-globaltool"
    fi
else
    echo ""
    echo "Tests failed!"
    exit 1
fi
