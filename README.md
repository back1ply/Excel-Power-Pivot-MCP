# Excel Power Pivot MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Excel Power Pivot data models. Create and manage DAX measures, relationships, and more through natural language.

## Features

- **Workbook Discovery & Connection** - Discover and connect to open Excel workbooks with Power Pivot data models
- **DAX Query Execution** - Run DAX queries and preview table data
- **Measure Management** - Create, update, and delete DAX measures with auto-formatting
- **Relationship Management** - Create, delete, and activate/deactivate table relationships
- **Model Exploration** - List tables, columns, measures, relationships, hierarchies, KPIs, and dependencies
- **Data Profiling** - Analyze column statistics (min, max, distinct count, nulls, sample values)
- **Power Query Discovery** - List Power Queries (M code) in the workbook
- **Table Management** - Add Excel tables to the data model, refresh tables/model

## Requirements

- **Windows 10/11** (required for Excel COM interop)
- **Microsoft Excel 2013+** with Power Pivot enabled

## Installation

### Option 1: Download Pre-built Binary (Recommended)

Download the latest release from the [Releases](https://github.com/yourusername/excel-powerpivot-mcp/releases) page.

### Option 2: Build from Source

```bash
git clone https://github.com/yourusername/excel-powerpivot-mcp.git
cd excel-powerpivot-mcp
dotnet build -c Release
```

The executable will be in `bin/Release/net8.0-windows/win-x64/ExcelPowerPivotMcp.exe`

## MCP Client Configuration

### Claude Desktop

Add to your `claude_desktop_config.json`:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel-powerpivot": {
      "command": "C:/path/to/ExcelPowerPivotMcp.exe"
    }
  }
}
```

### Cursor

Add to your Cursor MCP settings:

```json
{
  "mcpServers": {
    "excel-powerpivot": {
      "command": "C:/path/to/ExcelPowerPivotMcp.exe"
    }
  }
}
```

### Antigravity (VS Code)

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "excel-powerpivot": {
      "type": "stdio",
      "command": "C:/path/to/ExcelPowerPivotMcp.exe"
    }
  }
}
```

## Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `MCP_MAX_QUERY_ROWS` | Maximum rows returned from DAX queries | `1000` |
| `MCP_QUERY_TIMEOUT_SECONDS` | DAX query timeout | `120` |
| `MCP_CONNECTION_TIMEOUT_MS` | Excel connection timeout | `10000` |
| `MCP_CONNECTION_RETRY_COUNT` | Retry attempts for connection | `3` |
| `MCP_RESOURCES_PATH` | Path to resources folder | `./Resources` |

## Usage

### 1. Open Excel with Power Pivot

Open your Excel workbook that contains a Power Pivot data model.

### 2. Connect via MCP

The AI assistant will first discover and connect to your workbook:

```
AI: Let me discover your Excel workbooks...
→ discover_workbooks()
→ Found: MyModel.xlsx (has Power Pivot)

AI: Connecting to your workbook...
→ connect_workbook(workbook_name: "MyModel.xlsx")
→ Connected!
```

### 3. Explore and Modify

```
AI: Let me see what's in your data model...
→ get_model_summary()
→ 5 tables, 12 measures, 6 relationships

AI: I'll create a new measure for you...
→ create_measure(
    tableName: "Sales",
    measureName: "Total Revenue",
    expression: "SUM(Sales[Amount])"
  )

AI: Don't forget to save!
→ save_workbook()
```

## Available Tools

### Connection
| Tool | Description |
|------|-------------|
| `discover_workbooks` | Find open Excel workbooks with Power Pivot models |
| `connect_workbook` | Connect to a specific workbook |
| `get_connection_status` | Check current connection status |
| `save_workbook` | Save the connected workbook |
| `refresh_model` | Refresh the entire data model |

### Model Metadata
| Tool | Description |
|------|-------------|
| `get_model_summary` | Comprehensive model overview |
| `list_tables` | List all tables with row counts |
| `list_columns` | List columns in a table |
| `list_measures` | List measures with expressions |
| `list_relationships` | Show table relationships |
| `list_hierarchies` | List user-defined hierarchies |
| `list_kpis` | List Key Performance Indicators |
| `get_dependencies` | Show calculation dependencies |
| `list_power_queries` | List Power Queries (M code) |
| `list_excel_tables` | List Excel tables (ListObjects) |

### DAX Queries
| Tool | Description |
|------|-------------|
| `run_dax` | Execute DAX queries or preview table data |
| `analyze_column` | Get column statistics and sample values |

### Measure CRUD
| Tool | Description |
|------|-------------|
| `create_measure` | Create a new DAX measure |
| `update_measure` | Update expression/name/description |
| `delete_measure` | Delete a measure |

### Relationship CRUD
| Tool | Description |
|------|-------------|
| `create_relationship` | Create a table relationship |
| `delete_relationship` | Delete a relationship |
| `set_relationship_active` | Activate/deactivate a relationship |

### Table Operations
| Tool | Description |
|------|-------------|
| `add_table_to_model` | Add Excel table to data model |
| `set_format_string` | Set formatting for measure/column |
| `refresh_table` | Refresh a single table |

## Performance Tips

### Fast Measure Creation

Use `autoFormat: false` to skip DAX formatting for faster measure creation (~1.5s savings):

```json
{
  "tableName": "Sales",
  "measureName": "Quick Measure",
  "expression": "SUM(Sales[Amount])",
  "autoFormat": false
}
```

## Limitations

The following features are **NOT supported** due to Excel COM API limitations:

| Feature | Status |
|---------|--------|
| Create/Update/Delete Calculated Columns | ❌ |
| Set Column Descriptions | ❌ |
| DEFINE COLUMN in DAX queries | ❌ |
| Calculation Groups | ❌ |
| Perspectives | ❌ |
| Row-Level Security | ❌ |

## Troubleshooting

| Error | Solution |
|-------|----------|
| "Excel is not running" | Open Excel with your workbook |
| "Workbook not found" | Ensure the workbook is open in Excel |
| "No data model" | Create a Power Pivot data model first |
| "Not connected" | Call `connect_workbook` first |

## License

MIT License - see [LICENSE](LICENSE) file.

## Contributing

Contributions welcome! Please open an issue or pull request.

## Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io/) by Anthropic
- [DAX Formatter](https://www.daxformatter.com/) by SQLBI
