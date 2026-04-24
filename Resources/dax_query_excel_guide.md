---
name: 'DAX Query Guide for Excel Power Pivot'
description: 'Guidelines for writing DAX queries in Excel Power Pivot'
uriTemplate: 'resource://dax_query_excel_guide'
---
# DAX Query Guide for Excel Power Pivot

## Best Practices

When writing DAX queries:

* Include comments for clarity (DAX comments use `//` not `--`)
* Always include an ORDER BY clause when returning multiple rows
* Use meaningful variable names to improve readability
* Define measures with fully qualified names in DEFINE blocks

## DAX Query Syntax Rules

### Query Structure

#### DEFINE Block

* Use DEFINE at the beginning if the query includes VAR or MEASURE definitions
* Only use a single DEFINE block per query
* Separate definitions with new lines (no commas or semicolons)

#### Measure Definitions

**Important Distinction:**

* `DEFINE MEASURE` **creates** (defines) a temporary measure that exists only for the duration of the query
* `EVALUATE` **runs** (executes) the measure and returns results

**Defining Measures:**

* When defining: ALWAYS fully qualify the measure name including its host table
  * Example: `DEFINE MEASURE 'TableName'[MeasureName] = ...`
  * The host table must exist in the semantic model
* When using: Refer to the measure by name only, without the table qualifier
  * Example: Use `[MeasureName]` in expressions like `CALCULATE([MeasureName], ...)`
  
**To permanently save a measure**, use `create_measure` instead of DEFINE MEASURE.

#### Excel Power Pivot Restrictions

* **DEFINE COLUMN is NOT supported** in Excel Power Pivot
* **User-Defined Functions (UDFs) are NOT supported**
* Only DEFINE VAR and DEFINE MEASURE work in Excel

#### Ordering Results

* ALWAYS include an ORDER BY clause when EVALUATE returns multiple rows
* Do not use the ORDERBY function to sort the final query result

### CALCULATE and CALCULATETABLE Filter Rules

Boolean filters in CALCULATE or CALCULATETABLE have important restrictions:

* Cannot directly use a measure or another CALCULATE function
  * Solution: Use a variable to store the result, then reference the variable
* Cannot reference columns from two different tables
* When using the IN operator, the table operand must be a table variable, not a table expression
* Do not assign a boolean filter to a VAR definition

### SUMMARIZECOLUMNS Function

**Purpose**: Build summary tables with groupby columns and measure-like extension columns

**Parameter Order** (all optional, but must follow this order if used):

1. Groupby columns (can be from one or multiple tables)
2. Filters
3. Measures or measure-like calculations

**Key Rules**:

* Use SUMMARIZECOLUMNS as the default for building summary tables with measures
* Do not use SUMMARIZECOLUMNS without measure-like extension columns
* Returns only rows where at least one measure value is not BLANK
* Allows ANY number of measure-like calculations of arbitrary complexity
* DO NOT use boolean filters with SUMMARIZECOLUMNS

**When to Use Alternatives**:

* If there are no measures or calculations, use SUMMARIZE instead

### SUMMARIZE Function

**Allowed Pattern**:

```dax
SUMMARIZE(<table expression>, <column1>, …, <columnN>)
```

**Critical Restrictions**:

* NEVER use SUMMARIZE with measure-like expressions
  * ❌ Incorrect: `SUMMARIZE(<table>, <column>, "expr1", <expr1>, …)`
  * ✅ Correct: Use SUMMARIZECOLUMNS for measure calculations
* `VALUES('Table'[Column])` is shorthand for `SUMMARIZE('Table', 'Table'[Column])`

### Time Intelligence Functions

**DATESINPERIOD Rolling Windows**:

* The negative period offset must precisely match the number of periods required
* Examples:
  * 12-month window: Use -12 (not -11)
  * 3-month window: Use -3 (not -2)
* This prevents off-by-one errors

**Maintaining Clear Date Context**:

* Always establish a valid date context for time intelligence calculations
* Methods:
  * Include groupby columns from the date table, OR
  * Apply filters on date columns
* Without date context, time intelligence functions cannot determine a "current date" reference

## DAX Query Examples for Excel Power Pivot

### Example 1: Basic Aggregation

**Scenario**: Get total sales by product category.

```dax
EVALUATE
    SUMMARIZECOLUMNS(
        'Product'[Category],
        "Total Sales", SUM('Sales'[Amount])
    )
ORDER BY 'Product'[Category] ASC
```

### Example 2: Using DEFINE MEASURE for Testing

**Scenario**: Test a new measure before creating it permanently.

```dax
DEFINE
    MEASURE 'Sales'[Profit Margin] = 
        DIVIDE(
            SUM('Sales'[Profit]),
            SUM('Sales'[Revenue]),
            0
        )
EVALUATE
    SUMMARIZECOLUMNS(
        'Product'[Category],
        "Margin", [Profit Margin]
    )
ORDER BY 'Product'[Category] ASC
```

### Example 3: Filtering with Variables

**Scenario**: Find products with sales above average.

```dax
DEFINE
    VAR _AvgSales = AVERAGEX(VALUES('Product'[ProductKey]), [Total Sales])
    VAR _SummaryTable = SUMMARIZECOLUMNS(
        'Product'[Name],
        "Sales", [Total Sales]
    )
EVALUATE
    FILTER(_SummaryTable, [Sales] > _AvgSales)
ORDER BY [Sales] DESC
```

### Example 4: Year-over-Year Comparison

**Scenario**: Compare current year sales to previous year.

```dax
DEFINE
    MEASURE 'Sales'[Sales LY] = 
        CALCULATE(
            [Total Sales],
            SAMEPERIODLASTYEAR('Calendar'[Date])
        )
    MEASURE 'Sales'[YoY Growth] = 
        DIVIDE(
            [Total Sales] - [Sales LY],
            [Sales LY],
            BLANK()
        )
EVALUATE
    SUMMARIZECOLUMNS(
        'Calendar'[Year],
        "Current Year", [Total Sales],
        "Prior Year", [Sales LY],
        "Growth %", [YoY Growth]
    )
ORDER BY 'Calendar'[Year] ASC
```

### Example 5: Top N Analysis

**Scenario**: Find top 10 customers by sales.

```dax
EVALUATE
    TOPN(
        10,
        SUMMARIZECOLUMNS(
            'Customer'[Name],
            "Total Sales", [Total Sales]
        ),
        [Total Sales], DESC
    )
ORDER BY [Total Sales] DESC
```

### Example 6: Running Totals

**Scenario**: Calculate cumulative sales by date.

```dax
DEFINE
    MEASURE 'Sales'[Running Total] = 
        CALCULATE(
            [Total Sales],
            FILTER(
                ALL('Calendar'[Date]),
                'Calendar'[Date] <= MAX('Calendar'[Date])
            )
        )
EVALUATE
    SUMMARIZECOLUMNS(
        'Calendar'[Date],
        "Daily Sales", [Total Sales],
        "Cumulative", [Running Total]
    )
ORDER BY 'Calendar'[Date] ASC
```

### Example 7: Percent of Total

**Scenario**: Calculate each category's share of total sales.

```dax
DEFINE
    MEASURE 'Sales'[% of Total] = 
        DIVIDE(
            [Total Sales],
            CALCULATE([Total Sales], ALL('Product'[Category])),
            0
        )
EVALUATE
    SUMMARIZECOLUMNS(
        'Product'[Category],
        "Sales", [Total Sales],
        "Share", [% of Total]
    )
ORDER BY [Sales] DESC
```

### Example 8: Distinct Count with Filters

**Scenario**: Count unique customers who purchased in each year.

```dax
DEFINE
    MEASURE 'Sales'[Customer Count] = DISTINCTCOUNT('Sales'[CustomerKey])
EVALUATE
    SUMMARIZECOLUMNS(
        'Calendar'[Year],
        "Unique Customers", [Customer Count]
    )
ORDER BY 'Calendar'[Year] ASC
```

### Example 9: Using TREATAS for Virtual Relationships

**Scenario**: Apply a filter from one table to another without a physical relationship.

```dax
DEFINE
    VAR _TopProducts = TOPN(5, VALUES('Product'[ProductKey]), [Total Sales], DESC)
EVALUATE
    CALCULATETABLE(
        SUMMARIZECOLUMNS(
            'Product'[Name],
            "Sales", [Total Sales]
        ),
        _TopProducts
    )
ORDER BY [Sales] DESC
```

### Example 10: Handling BLANK Values

**Scenario**: Show products with no sales (BLANK handling).

```dax
DEFINE
    MEASURE 'Product'[Has Sales] = IF(ISBLANK([Total Sales]), "No", "Yes")
EVALUATE
    SUMMARIZECOLUMNS(
        'Product'[Name],
        "Sales Status", [Has Sales],
        "Amount", [Total Sales] + 0  // Adding 0 converts BLANK to 0
    )
ORDER BY 'Product'[Name] ASC
```

## Key Takeaways

1. **Always use ORDER BY** when returning multiple rows
2. **Store measure results in variables** before using in boolean filters
3. **Use DEFINE MEASURE for testing** before creating permanent measures
4. **DEFINE COLUMN is not supported** in Excel Power Pivot
5. **SUMMARIZECOLUMNS for measures**, SUMMARIZE for distinct values
6. **Establish date context** for time intelligence
7. **Max 1000 rows** returned per query

## See Also

* **[Common Workflows](resource://common_workflows)** - Quick reference for common Power Pivot operations
* **[Power Pivot Instructions](resource://excel_powerpivot_instructions)** - Getting started guide
* **[Measure Best Practices](resource://powerpivot_measure_best_practices)** - Writing effective DAX measures
