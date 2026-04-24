# Excel Power Pivot Performance Optimization Guide

Strategies for optimizing Power Pivot data models. Use `analyze_storage_statistics` and `analyze_vertipaq_compression` to diagnose issues.

## Table of Contents

1. [Understanding Vertipaq Compression](#understanding-vertipaq-compression)
2. [Data Type Optimization](#data-type-optimization)
3. [Column Design Best Practices](#column-design-best-practices)
4. [Cardinality Management](#cardinality-management)
5. [Relationship Optimization](#relationship-optimization)
6. [Table Design Patterns](#table-design-patterns)
7. [DAX Performance](#dax-performance)
8. [Diagnostic Workflow](#diagnostic-workflow)
9. [Common Anti-Patterns](#common-anti-patterns)

---

## Understanding Vertipaq Compression

Power Pivot uses **Vertipaq** columnar storage with two encoding methods:

1. **Value Encoding**: Dictionary-based. Best for low-cardinality (repeated values like Country, Status)
2. **Hash Encoding**: Hash-based. Best for high-cardinality (unique values like Transaction IDs)

### Compression Metrics

- **Dictionary Size**: Memory for unique values dictionary
- **Data Size**: Memory for compressed column data
- **Compression Ratio**: `DictionarySize / DataSize` — lower is better (<0.5 good, >2.0 poor)

---

## Data Type Optimization

### Rule 1: Use the Smallest Sufficient Data Type

**Impact**: Data type directly affects compression and memory usage.

| Data Type | Size per Value | When to Use |
| --------- | -------------- | ----------- |
| **Integer** | 1-8 bytes | Whole numbers, IDs, counts |
| **Decimal** | 8 bytes | Currency, precise calculations |
| **Float** | 8 bytes | Scientific data (avoid for finance) |
| **Date** | 8 bytes | Dates only (no time component) |
| **DateTime** | 8 bytes | Dates with time |
| **Text** | Variable | Strings (worst compression for high cardinality) |
| **Boolean** | 1 byte | True/False flags |

### Action: Convert Text to Numbers Where Possible

**❌ Bad**: Storing year as text

```dax
-- Column in Excel: "2024", "2023", "2022" (Text)
-- Memory: ~20 bytes per value + dictionary overhead
```

**✅ Good**: Store as integer

```dax
-- Column in Excel: 2024, 2023, 2022 (Whole Number)
-- Memory: ~4 bytes per value, perfect compression
```

**How to Check**:

```json
{
  "tool": "analyze_vertipaq_compression",
  "arguments": { "tableName": "Sales" }
}
```

Look for text columns with low cardinality that could be integers.

### Action: Remove Time from Dates If Not Needed

❌ Bad: `2024-01-15 00:00:00` → ✅ Good: `2024-01-15`

**Savings**: ~30-40% for date columns. Fix in Power Query: `= Date.From([DateTimeColumn])`

---

## Column Design Best Practices

### Rule 2: Avoid High-Cardinality String Columns

**Problem**: Unique strings compress poorly and consume massive memory.

**❌ Anti-Pattern**: Full text descriptions in fact tables

```text
OrderID | CustomerName          | ProductDescription
--------|----------------------|----------------------------
1       | "John Smith Corp."   | "Blue Widget Model XJ-2024"
2       | "Jane Doe LLC"       | "Red Widget Model XJ-2024"
```

**Impact**:

- Each unique string stored in dictionary
- High cardinality = large dictionary = poor compression

**✅ Solution 1**: Use IDs and dimension tables

```text
-- Sales Table (Fact)
OrderID | CustomerID | ProductID
--------|-----------|----------
1       | 1001      | 2024
2       | 1002      | 2024

-- Customer Table (Dimension)
CustomerID | CustomerName
-----------|-------------------
1001       | "John Smith Corp."
1002       | "Jane Doe LLC"
```

**Benefits**:

- CustomerID compresses perfectly (integers)
- Descriptions stored once in dimension table
- Memory reduced by 70-90%

**✅ Solution 2**: Use calculated columns (if needed for display only)

```dax
CustomerDisplay =
RELATED(Customers[CustomerName])
```

**Benefits**:

- Not stored in Vertipaq (calculated on-the-fly)
- Zero memory overhead
- Only works for display in visuals, not in measures

### Rule 3: Limit Decimal Precision

More decimal places = more unique values = worse compression.

Round to required precision in Power Query:

```m
= Table.TransformColumns(Source, {{"Price", each Number.Round(_, 2), type number}})
```

---

## Cardinality Management

### Rule 4: Understand Cardinality Impact

**Cardinality** = Number of unique values in a column

| Cardinality Level | Example | Compression | Optimization |
| ----------------- | ------- | ----------- | ------------ |
| **Very Low** (< 100) | Country, Status, Category | Excellent | Keep as-is |
| **Low** (100-1,000) | Product ID, Store ID | Good | Consider dimension tables |
| **Medium** (1K-100K) | Customer ID, SKU | Moderate | Use integers, not text |
| **High** (100K-1M) | Transaction ID, Order Number | Poor (if text) | **Must** use integers |
| **Very High** (> 1M) | Full descriptions, comments | Very Poor | Remove or move to dimension |

### Action: Identify High-Cardinality Columns

**Diagnostic Query**:

```json
{
  "tool": "get_storage_statistics",
  "arguments": {
    "level": "columns",
    "table_name": "Sales"
  }
}
```

**Red flags**:

- Cardinality > 100,000 for text columns
- Dictionary size > 10% of total table size
- Encoding = "HASH" for non-ID columns (indicates very high cardinality)

### Action: Remove Unnecessary High-Cardinality Columns

**Common culprits**:

- Full customer addresses (use AddressID instead)
- Product descriptions (use ProductID + dimension table)
- Transaction notes/comments (consider removing entirely)
- Concatenated keys (split into separate columns)

**Rule of thumb**: If you don't use it in a measure or filter, **remove it**.

---

## Relationship Optimization

### Rule 5: Use Integer Keys for Relationships

**❌ Bad**: Text-based relationships

```text
Sales.CustomerName → Customers.CustomerName
```

**Problems**:

- Poor compression (text keys)
- Slower relationship filtering
- Higher memory usage

**✅ Good**: Integer-based relationships

```text
Sales.CustomerID → Customers.CustomerID
```

**Benefits**:

- Perfect compression (integers)
- Fast relationship filtering
- Minimal memory overhead

### Rule 6: Minimize Relationship Count

**Problem**: Each relationship consumes memory for indexing.

**Best practices**:

- Remove unused relationships (mark inactive if needed for alternate paths)
- Avoid many-to-many relationships if possible
- Use single-direction filtering (not bi-directional) unless required

### Rule 7: Check Referential Integrity Violations

If `riViolations > 0` in storage stats:

- Orphaned rows exist (FK with no matching PK)
- Causes memory overhead and slower filtering
- Fix by cleaning data or removing orphaned rows

---

## Table Design Patterns

### Rule 8: Star Schema Over Flat Tables

**❌ Anti-Pattern**: Denormalized flat table

```text
Sales Table (5M rows × 20 columns)
- OrderID, CustomerName, CustomerCity, CustomerCountry
- ProductName, ProductCategory, ProductBrand
- OrderDate, ShipDate, Quantity, Revenue
```

**Problems**:

- Repeated customer/product data in every row
- Poor compression (duplicated strings)
- Large memory footprint

**✅ Best Practice**: Star schema

```text
Sales Fact (5M rows × 6 columns)
- OrderID, CustomerID, ProductID, OrderDate, Quantity, Revenue

Customers Dimension (10K rows × 5 columns)
- CustomerID, CustomerName, City, Country

Products Dimension (1K rows × 4 columns)
- ProductID, ProductName, Category, Brand
```

**Benefits**:

- Data stored once (in dimensions)
- Fact table uses integers (perfect compression)
- Memory reduced by 60-80%

### Rule 9: Use Date Tables Efficiently

**✅ Best Practice**: Single consolidated Date table

```dax
Date =
CALENDAR(
    DATE(2020, 1, 1),
    DATE(2030, 12, 31)
)
```

**Add calculated columns**:

- `Year = YEAR([Date])`
- `MonthNumber = MONTH([Date])`
- `MonthName = FORMAT([Date], "MMM")` → Sort by MonthNumber using "Sort by Column"
- `Quarter = "Q" & FORMAT([Date], "Q")`

**Memory impact**: Minimal (~4,000 rows for 11 years)

❌ Avoid separate date tables per fact table.

---

## DAX Performance

### Rule 10: Prefer Calculated Columns Over Imported Columns (Sometimes)

**When to use Calculated Columns**:

- ✅ Simple lookups: `= RELATED(Dimension[Column])`
- ✅ Date/time extraction: `= YEAR([OrderDate])`
- ✅ Simple arithmetic: `= [Quantity] * [UnitPrice]`

**Why**:

- Not stored in Vertipaq (calculated during query)
- Zero memory overhead
- Perfect for display-only columns

**When to use Imported Columns**:

- ✅ Complex transformations (better in Power Query)
- ✅ Data cleansing operations
- ✅ Columns used in relationships
- ✅ Columns filtered frequently

### Rule 11: Optimize Measure Performance

**❌ Slow measures**:

- Complex iterators (`SUMX`, `FILTER`) over large tables
- Nested `CALCULATE` with many filters
- String concatenation in measures

**✅ Fast measures**:

- Simple aggregations (`SUM`, `COUNT`, `AVERAGE`)
- Pre-aggregated tables (if appropriate)
- Variables to avoid recalculation

**Example optimization**:

```dax
-- ❌ Slow: Calculates line total for each row
Total Revenue =
SUMX(
    Sales,
    Sales[Quantity] * RELATED(Products[UnitPrice])
)

-- ✅ Fast: Pre-calculate line total in Sales table
-- Column: LineTotal = [Quantity] * RELATED(Products[UnitPrice])
Total Revenue = SUM(Sales[LineTotal])
```

---

## Diagnostic Workflow

### Step 1: Identify Memory Hotspots

**Run table-level analysis**:

```json
{
  "tool": "get_storage_statistics",
  "arguments": { "level": "tables" }
}
```

**Look for**:

- Tables with `bytesPerRow > 100` (indicates poor compression)
- Tables with `riViolations > 0` (data quality issues)
- Tables using disproportionate memory (>30% of total)

### Step 2: Drill into Problematic Tables

**Run column-level analysis**:

```json
{
  "tool": "analyze_vertipaq_compression",
  "arguments": { "tableName": "Sales" }
}
```

**Focus on columns with**:

- `percentOfTableSize > 10%` (single column using >10% of table memory)
- `encoding = "HASH"` for non-ID columns (high cardinality warning)
- `cardinality > 100,000` for text columns (compression failure)

### Step 3: Apply Targeted Optimizations

**Priority order**:

1. **Remove unused columns** (zero cost, immediate impact)
2. **Convert text to integers** (especially IDs and codes)
3. **Split into dimension tables** (for repeated data)
4. **Reduce decimal precision** (if applicable)
5. **Remove time from dates** (if time not needed)

### Step 4: Measure Impact

**Before optimization**:

```json
{
  "tool": "get_storage_statistics",
  "arguments": { "level": "tables" }
}
// Note: totalDataSize = 150,000,000 bytes (~150 MB)
```

**After optimization**:

```json
{
  "tool": "get_storage_statistics",
  "arguments": { "level": "tables" }
}
// Note: totalDataSize = 45,000,000 bytes (~45 MB)
// Result: 70% reduction
```

---

## Common Anti-Patterns

### Anti-Pattern 1: Importing "Everything" from Source

**❌ Problem**: Import all 100 columns from source system, use only 10

**Impact**:

- 90 unused columns consuming memory
- Slower refresh times
- Poor compression from irrelevant data

**✅ Solution**: Only import needed columns in Power Query

```m
// Remove unused columns in Power Query
= Table.SelectColumns(
    Source,
    {"OrderID", "CustomerID", "ProductID", "OrderDate", "Quantity", "Revenue"}
)
```

### Anti-Pattern 2: Storing Aggregates in Fact Tables

**❌ Problem**: Pre-calculating totals in imported data

```text
OrderID | Quantity | UnitPrice | LineTotal | OrderTotal
```

**Why bad**:

- `LineTotal` is redundant (can be calculated: `Quantity * UnitPrice`)
- `OrderTotal` is redundant (can be calculated in DAX)
- Wastes memory on derived data

**✅ Solution**: Calculate in DAX measures

```dax
Line Total = SUM(Sales[Quantity]) * SUM(Sales[UnitPrice])
Order Total = SUMX(Sales, [Line Total])
```

### Anti-Pattern 3: Using DateTime for Date-Only Data

**❌ Problem**: `OrderDate = 2024-01-15 00:00:00` (time is always midnight)

**Impact**:

- Wasted precision (time component unnecessary)
- Worse compression (more unique values due to second/millisecond variations)
- Slower date filtering

**✅ Solution**: Convert to Date in Power Query

```m
= Table.TransformColumnTypes(
    Source,
    {{"OrderDate", type date}}
)
```

### Anti-Pattern 4: Many Small Tables Instead of One Large Table

**❌ Problem**: Separate tables for each year/region/product type

**Why bad**:

- Overhead from table metadata
- Difficult to write DAX across tables
- Slower refresh (multiple connections)

**✅ Solution**: Single table with filter columns

```text
Sales Table
- OrderID, CustomerID, ProductID, OrderDate, Region, Year, Quantity, Revenue
```

**Use slicers/filters** instead of separate tables.

---

## Performance Targets

### Memory Benchmarks (Per Million Rows)

| Table Type | Target Size | Warning Threshold |
| ---------- | ----------- | ----------------- |
| **Optimized Fact Table** | 10-30 MB/M rows | > 50 MB/M rows |
| **Dimension Table** | 1-5 MB/M rows | > 10 MB/M rows |
| **Date Table** | < 1 MB (total) | > 2 MB |

### Compression Benchmarks

| Metric | Target | Action Needed |
| ------ | ------ | ------------- |
| **Bytes per Row** | < 50 bytes | Investigate if > 100 bytes |
| **Compression Ratio** | < 0.5 | Review columns if > 2.0 |
| **Dictionary Size** | < 5% of table | Optimize if > 10% |

### Query Performance

| Operation | Target | Warning |
| --------- | ------ | ------- |
| **Simple measure** | < 100ms | > 500ms (optimize DAX) |
| **Slicer click** | < 200ms | > 1s (check relationships) |
| **Model refresh** | < 5 min/GB | > 10 min/GB (check data types) |

---

## Optimization Checklist

Use this checklist after loading or modifying a Power Pivot model:

- [ ] Run `get_storage_statistics` to identify largest tables
- [ ] Run `analyze_vertipaq_compression` on top 3 largest tables
- [ ] Remove unused columns (check with `list_columns`)
- [ ] Convert text IDs to integers (check DataType in storage stats)
- [ ] Verify all relationships use integer keys (check with `list_relationships`)
- [ ] Check for RI violations (`riViolations = 0` in table stats)
- [ ] Remove time from date columns (if time not needed)
- [ ] Split repeated data into dimension tables (if high duplication)
- [ ] Review high-cardinality columns (cardinality > 100K)
- [ ] Validate compression ratios (< 0.5 for most columns)

---

## Getting Help

### Diagnostic Commands

**Full model assessment**:

```json
{
  "tool": "get_model_summary",
  "arguments": { "detail_level": "full" }
}
```

**Table memory breakdown**:

```json
{
  "tool": "get_storage_statistics",
  "arguments": { "level": "tables" }
}
```

**Column-level compression analysis**:

```json
{
  "tool": "analyze_vertipaq_compression",
  "arguments": { "tableName": "Sales" }
}
```

**Find dependencies before removing columns**:

```json
{
  "tool": "get_dependencies",
  "arguments": { "object_name": "CustomerName" }
}
```

### Additional Resources

- **SQLBI Vertipaq Analyzer**: <https://www.sqlbi.com/tools/vertipaq-analyzer/>
- **DAX Patterns**: <https://www.daxpatterns.com/>
- **Power Pivot Best Practices**: <https://www.sqlbi.com/articles/>

---

## Summary: Top 5 Optimizations

1. **Remove unused columns** → Immediate 20-40% memory reduction
2. **Convert text IDs to integers** → 50-70% reduction on ID columns
3. **Split into star schema** → 60-80% reduction for denormalized tables
4. **Remove time from dates** → 30-40% reduction on date columns
5. **Check for RI violations** → Fix data quality, improve performance

**Expected overall impact**: 50-80% memory reduction for poorly optimized models.

---

**Last Updated**: 2026-01-02
**MCP Server Version**: 2.7.0+
