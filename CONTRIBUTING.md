# Contributing

## Prerequisites

- Windows 10/11
- .NET 8 SDK
- Microsoft Excel with Power Pivot (Office 2016+ or Microsoft 365)
- An open Excel workbook with a Power Pivot data model (for manual testing)

## Build

```bash
# Debug
dotnet build

# Release (single-file exe)
dotnet publish -c Release
# Output: bin/Release/net8.0-windows/win-x64/publish/ExcelPowerPivotMcp.exe
```

## Test

```bash
# All tests
dotnet test

# Verbose output
dotnet test --logger "console;verbosity=detailed"

# Specific test class
dotnet test --filter "FullyQualifiedName~DaxHelpersTests"

# With coverage
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover /p:CoverletOutput=./TestResults/coverage
```

Most tests are unit tests with mocked dependencies and do not require Excel to be open.

## Manual Testing

With Excel open and a workbook containing a Power Pivot model:

```bash
# Run the MCP server directly
./bin/Release/net8.0-windows/win-x64/publish/ExcelPowerPivotMcp.exe

# Or use the test driver to simulate MCP client calls
python test_driver.py
```

## Project Structure

```
Tools/           # MCP tool definitions (presentation layer)
Core/Services/   # Business logic
Infrastructure/  # Excel COM interop and ADOMD.NET
PowerPivot/      # Connection management and COM helpers
Common/          # DTOs and utilities
Resources/       # Embedded markdown documentation
Tests/           # Unit tests
```

## Adding a New Tool

1. Create a class in `Tools/`
2. Use the `[McpTool]` attribute — auto-discovered via reflection
3. Inject services via constructor
4. Dispatch all Excel COM calls via `_excelDispatcher.InvokeAsync<T>()`
5. Release COM objects with `ComObjectManager.Release()`

## Adding a New Service

1. Define interface in `Core/Services/Interfaces.cs` or `Infrastructure/Services/IInfrastructureServices.cs`
2. Implement the class
3. Register as singleton in `Program.cs`

## COM Interop Rules

- **Always** dispatch to the STA thread via `ExcelStaService.InvokeAsync<T>()`
- **Always** release COM objects immediately after use
- Check `PowerPivotConnection.IsFatalComError()` on `COMException` — call `PowerPivotConnectionManager.Reset()` if true

## Pull Requests

- Keep PRs focused — one concern per PR
- Add or update tests for any changed behavior
- `dotnet test` must pass
- Follow existing code style (no comments unless the WHY is non-obvious)
