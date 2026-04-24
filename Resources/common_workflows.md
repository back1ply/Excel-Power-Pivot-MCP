---
name: 'Common Power Pivot Workflows'
description: 'Step-by-step workflows for common Power Pivot tasks'
uriTemplate: 'resource://common_workflows'
---

# Common Power Pivot Workflows

Quick reference for common operations in Excel Power Pivot via MCP.

## Connect to a Workbook

```text
1. discover_workbooks()           → See open workbooks with data models
2. connect_workbook(workbook_name: "Sales.xlsx")  → Connect to one
3. get_connection_status()        → Verify connected
```

## Create a Measure

```text
1. connect_workbook(...)          → Connect if not already
2. list_tables()                  → See available tables and columns
3. create_measure(
     table_name: "Sales",
     measure_name: "Total Revenue",
     expression: "SUM(Sales[Amount])"
   )
4. save_workbook()                → REQUIRED to persist!
```

## Test Before Creating

Use `run_dax` with DEFINE MEASURE to test an expression:

```dax
DEFINE 
    MEASURE 'Sales'[Test] = SUM(Sales[Amount])
EVALUATE 
    ROW("Result", [Test])
```

If it works, create the permanent measure.

## Update a Measure

```text
1. `list_measures()`                 → Find current expression
2. update_measure(
     measure_name: "Total Revenue",
     new_expression: "SUMX(Sales, Sales[Qty] * Sales[Price])"
   )
3. save_workbook()
```

## Calculated Columns

**Note**: Creating, updating, or deleting calculated columns is NOT supported via the Excel COM API. Calculated columns must be created manually in the Power Pivot window within Excel.

## Create a Relationship

```text
1. list_tables()                   → Find key columns
2. list_relationships()            → Check existing relationships
3. create_relationship(
     foreign_table: "Sales",
     foreign_column: "ProductKey",
     primary_table: "Products", 
     primary_column: "ProductKey",
     confirm: true                  → Required for safety
   )
4. save_workbook()
```

## Explore the Model

```text
1. `list_tables()`                   → All tables with columns
2. `list_measures()`                 → All measures with expressions
3. `list_relationships()`            → Model structure
```

## Common Patterns

### Ratio with DIVIDE

```dax
DIVIDE([Numerator], [Denominator], 0)
```

Always use DIVIDE, not `/`, to handle division by zero.

### Year-over-Year

```dax
VAR CurrentYear = [Total Sales]
VAR PriorYear = CALCULATE([Total Sales], SAMEPERIODLASTYEAR('Calendar'[Date]))
RETURN DIVIDE(CurrentYear - PriorYear, PriorYear, BLANK())
```

### Percent of Total

```dax
DIVIDE([Sales], CALCULATE([Sales], ALL('Products')), 0)
```

## See Also

- **[Power Pivot Instructions](resource://excel_powerpivot_instructions)** - Getting started with Excel Power Pivot
- **[DAX Query Guide](resource://dax_query_excel_guide)** - Complete DAX query reference for Excel
- **[Measure Best Practices](resource://powerpivot_measure_best_practices)** - Writing effective DAX measures
