# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Power Pivot MCP Server - A Model Context Protocol (MCP) server that enables AI assistants to interact with Excel Power Pivot data models. Written in C# .NET 8, this application bridges the gap between AI assistants and Excel Power Pivot through COM interop, enabling DAX query execution, measure management, relationship creation, and model exploration.

**Platform**: Windows-only (.NET 8 with Windows Forms for STA threading)
**Target**: Single-file executable (`ExcelPowerPivotMcp.exe`)

## Build and Test Commands

### Build
```bash
# Debug build
dotnet build

# Release build (single-file executable)
dotnet publish -c Release

# Output location: bin/Release/net8.0-windows/win-x64/publish/ExcelPowerPivotMcp.exe
```

### Test
```bash
# Run all tests
dotnet test

# Run tests with verbose output
dotnet test --logger "console;verbosity=detailed"

# Run specific test
dotnet test --filter "FullyQualifiedName~DaxHelpersTests"

# Run tests with code coverage
./run-tests-with-coverage.ps1  # Windows
./run-tests-with-coverage.sh   # Linux/Mac

# Manual coverage command
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover /p:CoverletOutput=./TestResults/coverage
```

### Development Testing
```bash
# Test the MCP server manually (requires open Excel workbook)
./bin/Release/net8.0-windows/win-x64/publish/ExcelPowerPivotMcp.exe

# Use test_driver.py to simulate MCP client interactions
python test_driver.py
```

## Architecture

### Layered Structure

```
┌─────────────────────────────────────┐
│  MCP Tools (Presentation Layer)    │  - Tools/, Prompts/, McpResourceProvider.cs
├─────────────────────────────────────┤
│  Core Services (Business Logic)     │  - Core/Services/
├─────────────────────────────────────┤
│  Infrastructure (External Systems)  │  - Infrastructure/Services/
├─────────────────────────────────────┤
│  PowerPivot (Connection Management) │  - PowerPivot/PowerPivotConnection.*.cs
├─────────────────────────────────────┤
│  Common (DTOs, Utilities)           │  - Common/
└─────────────────────────────────────┘
```

### Key Components

**Entry Point**: `Program.cs`
- Dependency injection configuration
- MCP server setup with stdio transport
- Auto-discovery of tools, resources, and prompts via reflection
- All services registered as singletons
- Logging configured to stderr only (stdout reserved for MCP JSON-RPC protocol)

**PowerPivot Connection**: `PowerPivot/PowerPivotConnection.cs` (partial class)
- Singleton pattern with resettable instance for COM error recovery
- Manages Excel Application COM objects and ADOMD.NET connections
- Split across multiple partial files for organization (Discovery, Measures, Relationships, Tables, Query)
- Critical: Uses `ComObjectManager` for proper COM object lifecycle management

**Threading Model**: Excel COM requires STA threading
- `ExcelStaService`: Dedicated STA thread with Windows Forms message pump
- All COM operations must be dispatched to this thread via `IExcelDispatcher.InvokeAsync<T>()`
- MCP server runs on thread pool, dispatches Excel operations to STA thread

**Infrastructure Services** (`Infrastructure/Services/`):
- `ExcelInteropService`: Excel.Application interactions
- `MeasureComService`: Measure create/update/delete operations
- `RelationshipComService`: Relationship management
- `TableComService`: Table operations - add, refresh, delete
- `AdoDmvService`: DMV metadata queries via ADOMD.NET
- `AdoDaxService`: DAX query execution via ADOMD.NET
- `ExcelStaService`: STA thread dispatcher

**Core Services** (`Core/Services/`):
- `MeasureService`: Measure CRUD with DAX formatting via SQLBI API
- `RelationshipService`: Relationship management
- `TableService`: Table operations (add to model, refresh)
- `ModelMetadataService`: Metadata aggregation and queries
- `DataProfileService`: Column profiling and statistics

**MCP Tools** (`Tools/`):
- `ConnectionTools.cs`: Workbook discovery, connection, save
- `DaxQueryTools.cs`: DAX execution and column analysis
- `MeasureCrudTools.cs`: Create, update, delete measures
- `ModelMetadataTools.cs`: List tables, columns, measures, relationships, hierarchies, KPIs
- `RelationshipCrudTools.cs`: Relationship CRUD operations
- `TableOperationTools.cs`: Add Excel tables to model, refresh tables

### Configuration

`McpConfiguration.cs` loads settings from environment variables with defaults:
- `MCP_MAX_QUERY_ROWS` (default: 1000)
- `MCP_QUERY_TIMEOUT_SECONDS` (default: 120)
- `MCP_VALIDATION_TIMEOUT_SECONDS` (default: 10)
- `MCP_DMV_TIMEOUT_SECONDS` (default: 30)
- `MCP_CONNECTION_RETRY_COUNT` (default: 3)
- `MCP_CONNECTION_RETRY_DELAY_MS` (default: 1000)
- `MCP_CONNECTION_TIMEOUT_MS` (default: 10000) - Increased from 5s to 10s to accommodate Excel COM initialization delays

## Critical Implementation Details

### COM Interop Rules

1. **Always use STA thread**: Dispatch Excel COM calls via `ExcelStaService.InvokeAsync<T>()`
2. **Proper COM cleanup**: Use `ComObjectManager.Release()` to release COM objects
3. **Fatal error detection**: `PowerPivotConnection.IsFatalComError()` identifies unrecoverable COM errors
4. **Connection recovery**: Call `PowerPivotConnectionManager.Reset()` on fatal errors

### Error Handling

- Tool-level exception filter wraps all tool executions, returning `IsError: true` with error message
- Connection validation detects stale Excel connections
- Polly retry policies for transient failures with exponential backoff
- **Duplicate object detection**: When attempting to create a measure that already exists, the error message now suggests using `update_measure` instead of `create_measure`
- **Connection warmup**: Initial connection includes a ~2 second warmup operation to ensure Excel data model is fully loaded, reducing timeout failures on first operations

### Resources

Markdown documentation embedded in assembly from `Resources/` folder:
- `excel_powerpivot_instructions.md`: Power Pivot guidelines
- `powerpivot_measure_best_practices.md`: DAX measure best practices
- `dax_query_excel_guide.md`: DAX query syntax for Excel
- `common_workflows.md`: Common usage patterns

MCP Resources expose runtime data via `model://` URIs.

## Excel Power Pivot Limitations

These features **DO NOT exist in Excel Power Pivot** (unlike Power BI):
- Calculation Groups
- Perspectives
- Row-Level Security (RLS)
- DEFINE COLUMN in DAX queries

These features **exist but cannot be managed via COM API**:
- Create/Update/Delete Calculated Columns (use Power Pivot window manually)
- Set Column Descriptions (use Power Pivot window manually)
- **Rename ModelTable objects** - The ModelTable.Name property is **read-only** for all connection types

### Table Naming Behavior

**Power Query Tables:**
- Table name is automatically set to match the Query name
- Cannot be renamed after creation (ModelTable.Name is read-only)
- If you rename the Query, the table is **replaced** (not renamed), breaking all measures and relationships
- **Solution**: Always create the Power Query with the desired table name upfront

**Linked Tables (Direct Excel Table Connections):**
- Table name is inherited from the Excel table name
- Cannot be renamed programmatically via COM API (ModelTable.Name is read-only)
- If you rename the Excel table, you must manually update the Power Pivot table through the Power Pivot window
- **No automatic synchronization** between Excel table name and Power Pivot table name after creation

**References:**
- [About the PowerPivot Model Object in Excel](https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-the-powerpivot-model-object-in-excel)
- [Change name of Power Pivot table created with Power Query](https://techcommunity.microsoft.com/t5/excel/change-name-of-power-pivot-table-created-with-power-query/td-p/2721270)
- [Add worksheet data to a Data Model using a linked table](https://support.microsoft.com/en-us/office/add-worksheet-data-to-a-data-model-using-a-linked-table-d3665fc3-99b0-479d-ba09-a37640f5be42)

### DMV Query Limitations

Excel Power Pivot's embedded ADO connection has **restricted access** to certain Dynamic Management Views (DMVs):

**❌ Unsupported DMVs** (removed from MCP server):
- `$SYSTEM.DISCOVER_STORAGE_TABLES` - Fails with error 0xC113000A
- `$SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS` - Fails with error 0xC113000A
- These Vertipaq storage queries are blocked/restricted in Excel Power Pivot's embedded connection

**✅ Supported DMVs** (working correctly):
- `$SYSTEM.MDSCHEMA_DIMENSIONS` - Table metadata
- `$SYSTEM.MDSCHEMA_MEASURES` - Measure metadata
- `$SYSTEM.MDSCHEMA_HIERARCHIES` - User-defined hierarchies
- `$SYSTEM.MDSCHEMA_LEVELS` - Column metadata
- `$SYSTEM.DISCOVER_CALC_DEPENDENCY` - Object dependencies
- `$SYSTEM.DBSCHEMA_CATALOGS` - Model compatibility level

**Alternatives for Vertipaq Analysis:**
- Use [DAX Studio](https://daxstudio.org/) to query storage DMVs directly
- Export the model to Power BI Desktop for comprehensive analysis tools

**References:**
- [Querying PowerPivot DMVs from Excel](https://blog.crossjoin.co.uk/2011/02/23/querying-powerpivot-dmvs-from-excel/)
- [Memory Analysis in PowerPivot](https://www.kasperonbi.com/what-is-eating-up-my-memory-powerpivot-excel-edition/)

## Development Guidelines

### When Adding New Tools

1. Create tool class in `Tools/` folder
2. Inherit from MCP SDK base classes and use `[McpTool]` attribute
3. Inject required services via constructor
4. Tool will be auto-discovered by `WithToolsFromAssembly()` in `Program.cs`

### When Adding New Services

1. Define interface in `Core/Services/Interfaces.cs` or `Infrastructure/Services/IInfrastructureServices.cs`
2. Implement service class
3. Register in `Program.cs` DI container as singleton
4. Inject dependencies via constructor

### When Working with Excel COM

- Always dispatch to STA thread via `_excelDispatcher.InvokeAsync<T>()`
- Release COM objects immediately after use with `ComObjectManager.Release()`
- Use `try-catch` for `COMException` and check `IsFatalComError()` for recovery

### DAX Formatting

- Measure creation calls SQLBI DAX Formatter API (daxformatter.com)
- Set `autoFormat: false` in measure creation to skip formatting (~1.5s faster)
- Formatter timeout is configured via `MCP_DAX_FORMATTER_TIMEOUT_SECONDS` (default: 10 seconds)

**Known Limitation**: The `Dax.Formatter` library creates its own HttpClient internally, which we cannot control. We mitigate this by:
- Using a shared singleton instance of `DaxFormatterClient` to reduce overhead
- Implementing aggressive caching with SHA256 keys to minimize API calls
- Wrapping calls with retry policies and graceful degradation

**Future Improvement**: Consider implementing direct HTTP calls to the DAX Formatter API to gain full control over HttpClient lifecycle via `IHttpClientFactory`.

## Release Process

The project uses GitHub CLI for manual releases:

```bash
# Build release
dotnet publish -c Release

# Create GitHub release manually (no automated workflow)
gh release create v2.7.5 --title "v2.7.5" --notes "Release notes here" ./bin/Publish/Standard/ExcelPowerPivotMcp.exe
```

Note: `.gitattributes` marks all files as `export-ignore`, making GitHub source archives empty by design (only release executable is distributed).
