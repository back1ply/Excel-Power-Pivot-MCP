---
name: 'Excel Power Pivot Instructions'
description: 'Guidelines for working with Excel Power Pivot data models via MCP'
uriTemplate: 'resource://excel_powerpivot_instructions'
---
# Excel Power Pivot MCP Guide

This MCP server enables AI assistants to interact with Power Pivot data models in Excel workbooks.

## Connection Workflow

### Step 1: Discover Workbooks

Before connecting, discover which Excel workbooks have Power Pivot data models:

```text
Tool: discover_workbooks
Arguments: {}
```

This returns a list of open workbooks with their paths and whether they have a data model.

### Step 2: Connect to a Workbook

Connect to a specific workbook by name or path:

```text
Tool: connect_workbook
Arguments: { "workbook_name": "MyWorkbook.xlsx" }
```

**Important**: The workbook must be open in Excel. You cannot connect to closed workbooks.

### Step 3: Verify Connection

Check the current connection status:

```text
Tool: get_connection_status
Arguments: {}
```

## Available Operations

### Read Operations (Always Safe)

`list_tables`, `list_columns`, `list_measures`, `list_relationships`, `get_model_summary`, `list_kpis`, `run_dax`, `list_excel_tables`, `list_power_queries`

### Write Operations (Require Save)

⚠️ Operations marked with † require `confirm=true` parameter.

`create_measure`, `update_measure`, `delete_measure`†, `create_relationship`†, `delete_relationship`†, `set_relationship_active`, `add_excel_table_to_model`†, `add_power_query_table_to_model`†, `refresh_table`, `refresh_model`

**Critical**: After write operations, use `save_workbook` to persist. Changes are in-memory until saved.

**Note**: Creating/updating/deleting **calculated columns** is NOT supported via Excel COM API.

## Excel Power Pivot Limitations

### Features NOT Supported in Excel Power Pivot (vs Power BI)

| Feature                        | Power BI  | Excel Power Pivot |
| ------------------------------ | --------- | ----------------- |
| DEFINE COLUMN in DAX queries   | ✅        | ❌                |
| User-Defined Functions (UDFs)  | ✅        | ❌                |
| Calculation Groups             | ✅        | ❌                |
| Perspectives                   | ✅        | ❌                |
| Translations                   | ✅        | ❌                |
| Row-Level Security (RLS)       | ✅        | ❌                |
| Object-Level Security          | ✅        | ❌                |
| Partitions Management          | ✅        | ❌                |
| Table Creation via DAX         | ✅        | ❌                |

### Features Supported via This MCP Server

| Feature                                            | Supported  |
| -------------------------------------------------- | ---------- |
| Measure Management (create/update/delete)          | ✅         |
| Relationship Management (create/delete/activate)   | ✅         |
| Add Excel Tables to Data Model                     | ✅         |
| DAX Query Execution                                | ✅         |
| Power Query Discovery                              | ✅         |
| Set Format Strings                                 | ✅         |
| Table/Model Refresh                                | ✅         |

**NOT Supported** (Excel COM API limitation):

| Feature                       | Status  |
| ----------------------------- | ------- |
| Calculated Column Management  | ❌      |
| Column Description Setting    | ❌      |

### Supported Excel Versions

- Excel 2013 and later (with Power Pivot add-in enabled)
- Excel 2016, 2019, 2021
- Microsoft 365 Excel

### Connection Architecture

Excel Power Pivot uses in-process VertiPaq (unlike Power BI's separate msmdsrv.exe process). This MCP server uses Excel COM automation, which is why:

- The workbook must be open in Excel
- Only one Excel instance can be accessed at a time
- Changes require saving the workbook

## DAX Query Execution

### Basic Query Structure

```dax
EVALUATE
    <table expression>
ORDER BY
    <column> [ASC|DESC]
```

### When to Use Just EVALUATE

Use `EVALUATE` alone when querying existing measures or doing simple aggregations:

```dax
EVALUATE
    SUMMARIZECOLUMNS(
        'Date'[Year],
        "Total", [Total Sales]   // existing model measure
    )
ORDER BY 'Date'[Year]
```

### When to Use DEFINE MEASURE + EVALUATE

Use `DEFINE MEASURE` when you want to **test a new measure** before creating it permanently:

```dax
DEFINE
    MEASURE 'TableName'[TempMeasure] = SUM('Sales'[Amount])
EVALUATE
    SUMMARIZECOLUMNS(
        'Date'[Year],
        "Total", [TempMeasure]
    )
ORDER BY 'Date'[Year]
```

**Why use DEFINE MEASURE?**

- Test new DAX logic without modifying the model
- If the query fails, no harm done - nothing was saved
- Temporarily override existing measures for "what if" analysis
- Once satisfied, use `create_measure` to save it permanently

**Note**: The host table in DEFINE MEASURE must exist in the model.

### Query Limits

- Maximum 1000 rows returned by default
- Use `max_rows` parameter to adjust (up to 1000)
- For larger datasets, use aggregations or filtering

## Measure Management

### Creating Measures

```text
Tool: create_measure
Arguments: {
    "table_name": "Sales",
    "measure_name": "Total Sales",
    "expression": "SUM(Sales[Amount])",
    "description": "Sum of all sales amounts",
    "autoFormat": true
}
```

**Best Practices**:

- Choose a logical host table (usually the fact table or a related table)
- Use clear, descriptive names
- Include descriptions for documentation
- Test the expression with `run_dax` using DEFINE MEASURE first
- Set `autoFormat=false` for faster measure creation (~1.5s savings)

### Performance Tip: autoFormat Parameter

Both `create_measure` and `update_measure` support an `autoFormat` parameter:

```text
Tool: create_measure
Arguments: {
    "table_name": "Sales",
    "measure_name": "Total Sales",
    "expression": "SUM(Sales[Amount])",
    "autoFormat": false
}
```

- `autoFormat=true` (default): Formats DAX for readability via dax.formatter API (~1.5s)
- `autoFormat=false`: Skips formatting for faster execution

### Updating Measures

```text
Tool: update_measure
Arguments: {
    "measure_name": "Total Sales",
    "new_expression": "SUMX(Sales, Sales[Quantity] * Sales[Price])",
    "new_description": "Calculated total from quantity and price",
    "autoFormat": true
}
```

### Deleting Measures

```text
Tool: delete_measure
Arguments: {
    "measure_name": "Total Sales",
    "confirm": true
}
```

**Warning**: This cannot be undone. The `confirm=true` parameter is required for safety.

## Calculated Columns

**Important Limitation**: The Excel COM API does NOT support programmatic creation, update, or deletion of calculated columns.

To create calculated columns:

1. Open the Power Pivot window in Excel (Power Pivot tab → Manage)
2. Navigate to the table
3. Add a calculated column manually

## Relationship Management

### Creating Relationships

```text
Tool: create_relationship
Arguments: {
    "foreign_table": "Sales",
    "foreign_column": "ProductKey",
    "primary_table": "Product",
    "primary_column": "ProductKey",
    "confirm": true
}
```

**Requirements**:

- One-to-Many relationship (Foreign Table -> Primary Table)
- Types must match between columns

### Deleting Relationships

```text
Tool: delete_relationship
Arguments: {
    "foreign_table": "Sales",
    "foreign_column": "ProductKey",
    "primary_table": "Product",
    "primary_column": "ProductKey",
    "confirm": true
}
```

## Table Management

### Listing Excel Tables

```text
Tool: list_excel_tables
Arguments: {}
```

This returns all Excel tables (ListObjects) in the workbook and indicates if they are already in the data model.

### Adding Tables to Model

```text
Tool: add_table_to_model
Arguments: {
    "table_name": "NewData",
    "use_power_query": true
}
```

**Options**:

- `use_power_query`: If true (recommended), adds via Power Query for better durability. If false, adds directly as a linked table.

## Error Handling

### Common Errors

| Error                  | Cause                             | Solution                                                        |
| ---------------------- | --------------------------------- | --------------------------------------------------------------- |
| "Excel is not running" | No Excel process found            | Open Excel with the workbook                                    |
| "Workbook not found"   | Workbook not open                 | Open the specific workbook in Excel                             |
| "No data model"        | Workbook has no Power Pivot model | Create a data model in the workbook first                       |
| "Not connected"        | No active connection              | Use `connect_workbook` first                                    |
| "Table not found"      | Invalid table name                | Check table names with `list_tables` (names are case-sensitive) |
| "Measure not found"    | Invalid measure name              | Check measure names with `list_measures`                        |

### Error Codes (Troubleshooting)

The MCP server translates common COM/OleDb error codes, but you may still encounter raw codes:

| Error Code   | Meaning                           | Suggestion                                                                                                                                                               |
| ------------ | --------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `0xC113000A` | Table not found or schema changed | Verify table names with `list_tables` and refresh the model via `refresh_model`. This often happens if you query a table immediately after adding it without refreshing. |
| `0x80040E14` | DAX Syntax Error                  | Check your DAX expression for typos, missing brackets, or invalid function usage. Use `list_columns` to verify column names.                                             |
| `0x80004005` | Unspecified Error                 | Ensure Excel is not in a dialog (like "Save As") or cell-editing mode (cursor inside a cell).                                                                            |

### DAX Query Errors

If a DAX query fails, the error message from the VertiPaq engine is returned. Common issues:

- Missing table/column references (check `list_tables` and `list_columns`)
- Syntax errors (check parentheses, commas)
- Invalid function usage
- Circular references

## See Also

- **[Common Workflows](resource://common_workflows)** - Step-by-step workflows for common Power Pivot tasks
- **[DAX Query Guide](resource://dax_query_excel_guide)** - Complete DAX query reference for Excel
- **[Measure Best Practices](resource://powerpivot_measure_best_practices)** - Writing effective DAX measures
