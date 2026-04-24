# Power Pivot DMV Coverage Reference

This document lists all DMV (Dynamic Management View) schema rowsets available in Excel Power Pivot and tracks MCP tool coverage.

**Compatibility Level:** 1103 (Excel 2013) or 1100 (Excel 2010)  
**Note:** TMSCHEMA_* views require compat level 1200+ and are **NOT available** in Excel Power Pivot. The COM API is used for relationships.

---

## DBSCHEMA_* (Database Schema) - 4 views

| DMV                       | Purpose                                 | MCP Tool            | Status                     |
| ------------------------- | --------------------------------------- | ------------------- | -------------------------- |
| `DBSCHEMA_CATALOGS`       | Model name, compat level, last modified | `get_model_summary` | [x] Implemented            |
| `DBSCHEMA_COLUMNS`        | Column metadata                         | —                   | [ ] Using MDSCHEMA instead |
| `DBSCHEMA_PROVIDER_TYPES` | Available data types                    | —                   | [ ] Low priority           |
| `DBSCHEMA_TABLES`         | Table list                              | —                   | [ ] Using MDSCHEMA instead |

---

## MDSCHEMA_* (Multidimensional Schema) - 14 views

| DMV                                | Purpose                 | MCP Tool           | Status                                 |
| ---------------------------------- | ----------------------- | ------------------ | -------------------------------------- |
| `MDSCHEMA_DIMENSIONS`              | Tables                  | `list_tables`      | [x] Implemented                        |
| `MDSCHEMA_LEVELS`                  | Columns                 | `list_columns`     | [x] Implemented                        |
| `MDSCHEMA_MEASURES`                | Measures                | `list_measures`    | [x] Implemented                        |
| `MDSCHEMA_KPIS`                    | KPIs                    | `list_kpis`        | [x] Implemented                        |
| `MDSCHEMA_CUBES`                   | Model/cube info         | —                  | [ ] Could add                          |
| `MDSCHEMA_HIERARCHIES`             | Hierarchies             | `list_hierarchies` | [x] Implemented                        |
| `MDSCHEMA_FUNCTIONS`               | DAX functions list      | —                  | [ ] Low priority                       |
| `MDSCHEMA_MEASUREGROUPS`           | Measure groups (tables) | —                  | [ ] Redundant                          |
| `MDSCHEMA_MEASUREGROUP_DIMENSIONS` | Table cardinality       | —                  | [ ] Table-level only, not column FK/PK |
| `MDSCHEMA_ACTIONS`                 | Drillthrough actions    | —                  | [ ] Rarely used                        |
| `MDSCHEMA_MEMBERS`                 | Dimension members       | —                  | [ ] Not relevant                       |
| `MDSCHEMA_PROPERTIES`              | Cell properties         | —                  | [ ] Not relevant                       |
| `MDSCHEMA_SETS`                    | Named sets              | —                  | [ ] Rarely used                        |
| `MDSCHEMA_INPUT_DATASOURCES`       | Data sources            | —                  | [ ] Could add                          |

---

## DISCOVER_* (Discovery) - 35 views

### High Value

| DMV                        | Purpose                              | MCP Tool           | Status          |
| -------------------------- | ------------------------------------ | ------------------ | --------------- |
| `DISCOVER_CALC_DEPENDENCY` | What measures/columns depend on what | `get_dependencies` | [x] Implemented |

### Storage & Performance

| DMV                                      | Purpose                | MCP Tool | Status                          |
| ---------------------------------------- | ---------------------- | -------- | ------------------------------- |
| `DISCOVER_STORAGE_TABLES`                | Table storage info     | —        | [ ] Could add for perf analysis |
| `DISCOVER_STORAGE_TABLE_COLUMNS`         | Column storage details | —        | [ ] Could add for perf analysis |
| `DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS` | Segment info           | —        | [ ] Advanced                    |
| `DISCOVER_OBJECT_MEMORY_USAGE`           | Memory by object       | —        | [ ] Advanced                    |
| `DISCOVER_MEMORYUSAGE`                   | Overall memory         | —        | [ ] Advanced                    |
| `DISCOVER_PARTITION_STAT`                | Partition stats        | —        | [ ] Advanced                    |
| `DISCOVER_DIMENSION_STAT`                | Table stats            | —        | [ ] Could add                   |
| `DISCOVER_PARTITION_DIMENSION_STAT`      | Partition/table stats  | —        | [ ] Advanced                    |

### Session & Connection (Monitoring)

| DMV                     | Purpose          | MCP Tool | Status          |
| ----------------------- | ---------------- | -------- | --------------- |
| `DISCOVER_SESSIONS`     | Active sessions  | —        | [ ] Monitoring  |
| `DISCOVER_CONNECTIONS`  | Connections      | —        | [ ] Monitoring  |
| `DISCOVER_COMMANDS`     | Running commands | —        | [ ] Monitoring  |
| `DISCOVER_TRANSACTIONS` | Transactions     | —        | [ ] Monitoring  |
| `DISCOVER_LOCKS`        | Locks            | —        | [ ] Monitoring  |
| `DISCOVER_JOBS`         | Background jobs  | —        | [ ] Monitoring  |

### Metadata & Config

| DMV                       | Purpose           | MCP Tool | Status         |
| ------------------------- | ----------------- | -------- | -------------- |
| `DISCOVER_PROPERTIES`     | Server properties | —        | [ ] Config     |
| `DISCOVER_KEYWORDS`       | Reserved keywords | —        | [ ] Reference  |
| `DISCOVER_LITERALS`       | Literal formats   | —        | [ ] Reference  |
| `DISCOVER_ENUMERATORS`    | Enumerations      | —        | [ ] Reference  |
| `DISCOVER_DATASOURCES`    | Data sources      | —        | [ ] Could add  |
| `DISCOVER_DB_CONNECTIONS` | DB connections    | —        | [ ] Could add  |
| `DISCOVER_CSDL_METADATA`  | CSDL metadata     | —        | [ ] Advanced   |
| `DISCOVER_XML_METADATA`   | Full XML metadata | —        | [ ] Advanced   |

### Tracing & XEvents

| DMV                                      | Purpose           | Status       |
| ---------------------------------------- | ----------------- | ------------ |
| `DISCOVER_TRACES`                        | Trace definitions | [ ] Advanced |
| `DISCOVER_TRACE_COLUMNS`                 | Trace columns     | [ ] Advanced |
| `DISCOVER_TRACE_EVENT_CATEGORIES`        | Event categories  | [ ] Advanced |
| `DISCOVER_TRACE_DEFINITION_PROVIDERINFO` | Provider info     | [ ] Advanced |
| `DISCOVER_XEVENT_*` (5 views)            | Extended events   | [ ] Advanced |

### Other

| DMV                             | Purpose          | Status            |
| ------------------------------- | ---------------- | ----------------- |
| `DISCOVER_SCHEMA_ROWSETS`       | List of all DMVs | [ ] Meta          |
| `DISCOVER_INSTANCES`            | Server instances | [ ] N/A for Excel |
| `DISCOVER_LOCATIONS`            | Server locations | [ ] N/A for Excel |
| `DISCOVER_MASTER_KEY`           | Encryption key   | [ ] N/A           |
| `DISCOVER_MEMORYGRANT`          | Memory grants    | [ ] Advanced      |
| `DISCOVER_PERFORMANCE_COUNTERS` | Perf counters    | [ ] Advanced      |
| `DISCOVER_RESOURCE_POOLS`       | Resource pools   | [ ] N/A for Excel |
| `DISCOVER_RING_BUFFERS`         | Ring buffers     | [ ] Advanced      |
| `DISCOVER_COMMAND_OBJECTS`      | Command objects  | [ ] Advanced      |
| `DISCOVER_OBJECT_ACTIVITY`      | Object activity  | [ ] Advanced      |

---

## DMSCHEMA_* (Data Mining) - 10 views

| DMV                                  | Purpose           | Status                           |
| ------------------------------------ | ----------------- | -------------------------------- |
| `DMSCHEMA_MINING_MODELS`             | Mining models     | [ ] Not relevant for Power Pivot |
| `DMSCHEMA_MINING_COLUMNS`            | Mining columns    | [ ] Not relevant                 |
| `DMSCHEMA_MINING_FUNCTIONS`          | Mining functions  | [ ] Not relevant                 |
| `DMSCHEMA_MINING_MODEL_CONTENT`      | Model content     | [ ] Not relevant                 |
| `DMSCHEMA_MINING_MODEL_CONTENT_PMML` | PMML format       | [ ] Not relevant                 |
| `DMSCHEMA_MINING_MODEL_XML`          | XML format        | [ ] Not relevant                 |
| `DMSCHEMA_MINING_SERVICES`           | Mining services   | [ ] Not relevant                 |
| `DMSCHEMA_MINING_SERVICE_PARAMETERS` | Service params    | [ ] Not relevant                 |
| `DMSCHEMA_MINING_STRUCTURES`         | Mining structures | [ ] Not relevant                 |
| `DMSCHEMA_MINING_STRUCTURE_COLUMNS`  | Structure columns | [ ] Not relevant                 |

---

## Coverage Summary

| Family       | Total  | Implemented | Relevant | Coverage |
| ------------ | ------ | ----------- | -------- | -------- |
| `DBSCHEMA_*` | 4      | **1**       | 1        | 100%     |
| `MDSCHEMA_*` | 14     | **5**       | 6        | 83%      |
| `DISCOVER_*` | 35     | **1**       | 5        | 20%      |
| `DMSCHEMA_*` | 10     | 0           | 0        | N/A      |
| **Total**    | **63** | **7**       | **12**   | **58%**  |

---

## Priority Additions

1. `DISCOVER_CALC_DEPENDENCY` -> `get_dependencies` (prevent breaking changes)
2. `DBSCHEMA_CATALOGS` -> Add to `get_model_summary` (compat level)
3. `MDSCHEMA_HIERARCHIES` -> `list_hierarchies` (if model uses them)
4. `DISCOVER_STORAGE_TABLES` -> Performance analysis

---

## Notes

- **TMSCHEMA_* views** require compat level 1200+ and are NOT available in Excel Power Pivot
- **Relationships** require COM API because MDSCHEMA only provides table-level cardinality, not column-level FK/PK details
- COM API is used for all write operations and DAX execution via Excel's internal ADO connection
