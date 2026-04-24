# Excel Power Pivot MCP - Development Guide

## Prerequisites

### Required Software

- **Windows 10/11** (required for Excel COM interop)
- **Microsoft Excel 2013+** with Power Pivot enabled
- **.NET 8.0 SDK** (or later) - [Download](https://dotnet.microsoft.com/download/dotnet/8.0)
- **Visual Studio 2022** or **VS Code** with C# extensions

### Enabling Power Pivot in Excel

1. Open Excel → File → Options → Add-ins
2. Manage: COM Add-ins → Go
3. Check "Microsoft Power Pivot for Excel"
4. Click OK

---

## Getting Started

### Clone and Build

```bash
# Clone the repository
git clone https://github.com/back1ply/Excel-Power-Pivot-MCP.git
cd Excel-Power-Pivot-MCP

# Restore dependencies
dotnet restore

# Build
dotnet build

# Run in development mode
dotnet run
```

### Quick Test

1. Open Excel with a workbook containing a Power Pivot data model
2. Run the MCP server: `dotnet run`
3. Connect via Claude Desktop, Cursor, or another MCP client

---

## Project Structure

```text
ExcelPowerPivotMcp.sln           # Solution file
├── ExcelPowerPivotMcp.csproj    # Main project
└── Tests/
    └── ExcelPowerPivotMcp.Tests.csproj  # Test project
```

---

## Build Commands

| Command | Description |
| ------- | ----------- |
| `dotnet build` | Build in Debug mode |
| `dotnet build -c Release` | Build in Release mode |
| `dotnet run` | Run in development mode |
| `dotnet test` | Run unit tests |
| `dotnet publish -c Release` | Create self-contained executable |

### PowerShell Scripts

```powershell
# Build script with additional options
.\build.ps1

# Release script for distribution
.\release.ps1
```

---

## Development Workflow

### Adding a New MCP Tool

1. Create or edit a file in `/Tools`
2. Add a public static method with `[McpServerTool]` attribute:

```csharp
[McpServerTool(Name = "my_new_tool")]
[Description("Description for the AI client")]
public static async Task<object> MyNewTool(
    [FromContainer] SomeService service,
    [Description("Parameter description")] string param1)
{
    // Implementation
    return new { result = "success" };
}
```

1. The tool is automatically discovered and registered at startup

### Adding a New Core Service

1. Define interface in `/Core/Services/Interfaces.cs`
2. Create implementation in `/Core/Services/`
3. Register in `Program.cs`:

```csharp
builder.Services.AddSingleton<IMyService, MyService>();
```

### Adding a New Resource

1. Add markdown file to `/Resources`
2. Update `McpResourceProvider.cs` to expose it with a `model://` URI

---

## Configuration

### Environment Variables

| Variable | Default | Description |
| -------- | ------- | ----------- |
| `MCP_SERVER_NAME` | `excel-powerpivot-mcp` | Server name in MCP |
| `MCP_MAX_QUERY_ROWS` | `1000` | Max rows returned from DAX |
| `MCP_QUERY_TIMEOUT_SECONDS` | `120` | DAX query timeout |
| `MCP_VALIDATION_TIMEOUT_SECONDS` | `10` | DAX validation timeout |
| `MCP_DMV_TIMEOUT_SECONDS` | `30` | DMV query timeout |
| `MCP_CONNECTION_RETRY_COUNT` | `3` | Connection retry attempts |
| `MCP_CONNECTION_RETRY_DELAY_MS` | `1000` | Base retry delay |
| `MCP_CONNECTION_TIMEOUT_MS` | `5000` | Single attempt timeout |

### Setting in Development

```bash
# PowerShell
$env:MCP_MAX_QUERY_ROWS = "5000"
dotnet run

# Command Prompt
set MCP_MAX_QUERY_ROWS=5000
dotnet run
```

---

## Testing

### Running Tests

```bash
# Run all tests
dotnet test

# Run with verbose output
dotnet test --logger "console;verbosity=detailed"

# Run specific test
dotnet test --filter "FullyQualifiedName~DaxHelpersTests"
```

### Manual Testing with Python Driver

A Python test driver is available for integration testing:

```bash
python test_driver.py
```

---

## Debugging

### Visual Studio

1. Open `ExcelPowerPivotMcp.sln`
2. Set `ExcelPowerPivotMcp` as startup project
3. Press F5 to debug

### VS Code

1. Open folder in VS Code
2. Install C# extension
3. Press F5 or use "Run and Debug"

### Logging

All logs go to `stderr` (stdout is reserved for MCP JSON-RPC):

```bash
# View logs
dotnet run 2>&1 | tee debug.log
```

---

## Publishing

### Single-File Executable

```bash
dotnet publish -c Release -o ./publish
```

This creates a self-contained `ExcelPowerPivotMcp.exe` (~35MB) that:

- Includes .NET runtime
- Embeds all resources
- Works without .NET installation on target machine

### Release Checklist

1. Update version in `.csproj`
2. Run tests: `dotnet test`
3. Build release: `.\release.ps1`
4. Test the published executable
5. Create GitHub release with executable

---

## Code Style

### Conventions

- Use file-scoped namespaces
- Prefer `var` for local variables when type is obvious
- Use `// NOTE:` comments for important implementation notes
- All service methods should be async

### Analyzers

The project uses:

- `AnalysisLevel: latest-recommended`
- `EnforceCodeStyleInBuild: true`

Warnings are treated as suggestions (not errors) except for critical issues.

---

## Common Issues

### "Excel is not running"

Make sure Excel is open with the target workbook before connecting.

### COM Exception on startup

Ensure you're running on Windows with Excel installed. The project cannot run on macOS/Linux.

### Connection lost during operation

Excel may have closed or the workbook was closed. Use `recover_connection` prompt or reconnect.
