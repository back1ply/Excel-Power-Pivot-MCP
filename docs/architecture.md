# Excel Power Pivot MCP - Architecture

## System Architecture

### Layered Architecture

The application follows Clean Architecture with distinct layers:

```text
┌─────────────────────────────────────────────────────────────┐
│                       Presentation                           │
│          (MCP Tools, Resources, Prompts)                     │
├─────────────────────────────────────────────────────────────┤
│                       Application                            │
│                    (Core Services)                           │
├─────────────────────────────────────────────────────────────┤
│                      Infrastructure                          │
│        (Excel COM, ADOMD.NET, External APIs)                 │
├─────────────────────────────────────────────────────────────┤
│                         Common                               │
│           (DTOs, Utilities, Shared Code)                     │
└─────────────────────────────────────────────────────────────┘
```

---

## Technology Stack

| Category | Technology | Version | Notes |
| -------- | ---------- | ------- | ----- |
| **Runtime** | .NET | 8.0 | Windows-specific (COM interop) |
| **MCP SDK** | ModelContextProtocol | 0.5.0-preview.1 | Official C# SDK |
| **Hosting** | Microsoft.Extensions.Hosting | 8.0.1 | DI container & lifecycle |
| **Resilience** | Polly | 8.2.0 | Retry policies |
| **DAX Formatting** | Dax.Formatter | 1.1.0 | SQLBI's DAX formatter API |
| **Platform** | Windows 10/11 | - | Required for Excel COM |
| **Excel** | Excel 2013+ | - | Power Pivot enabled |

---

## Component Architecture

### 1. Entry Point (`Program.cs`)

The application entry point:

- Configures dependency injection container
- Registers all services as singletons
- Sets up MCP server with stdio transport
- Configures logging (stderr only - stdout reserved for MCP protocol)
- Auto-discovers tools, resources, and prompts via reflection

### 2. Core Services Layer

Business logic services implementing domain operations:

| Interface | Implementation | Responsibility |
| ----------- | ---------------- | ---------------- |
| `IMeasureService` | `MeasureService` | Measure CRUD operations |
| `IRelationshipService` | `RelationshipService` | Relationship management |
| `ITableService` | `TableService` | Table operations |
| `IModelMetadataService` | `ModelMetadataService` | Metadata queries |
| `IDataProfileService` | `DataProfileService` | Data profiling |

### 3. Infrastructure Services Layer

External integrations and platform-specific implementations:

| Interface | Implementation | Responsibility |
| ----------- | ---------------- | ---------------- |
| `IExcelInteropService` | `ExcelInteropService` | Excel Application interactions |
| `IPowerPivotComService` | `PowerPivotComService` | Power Pivot COM model |
| `IDmvService` | `AdoDmvService` | DMV metadata queries |
| `IDaxService` | `AdoDaxService` | DAX query execution |
| `IExcelDispatcher` | `ExcelStaService` | STA thread dispatching |
| `IInMemoryLogReader` | `InMemoryLoggerProvider` | Log access for MCP resource |

### 4. PowerPivot Layer

Connection management to Excel Power Pivot:

```text
PowerPivotConnection (partial class)
├── PowerPivotConnection.cs       # Core connection logic
├── PowerPivotConnection.Discovery.cs  # Workbook discovery
└── [Other partial files for organization]
```

Key responsibilities:

- Excel COM object management
- ADOMD.NET connection handling
- Connection validation and recovery
- Dirty state tracking for unsaved changes

### 5. MCP Tools Layer

Auto-discovered tool handlers in `/Tools`:

| File | Tool Categories |
| ------ | ----------------- |
| `ConnectionTools.cs` | `discover_workbooks`, `connect_workbook`, `get_connection_status`, `save_workbook`, `refresh_model` |
| `DaxQueryTools.cs` | `run_dax`, `analyze_column` |
| `MeasureCrudTools.cs` | `create_measure`, `update_measure`, `delete_measure` |
| `ModelMetadataTools.cs` | `get_model_summary`, `list_tables`, `list_columns`, `list_measures`, `list_relationships`, `list_hierarchies`, `list_kpis`, `get_dependencies`, `list_power_queries`, `list_excel_tables` |
| `RelationshipCrudTools.cs` | `create_relationship`, `delete_relationship`, `set_relationship_active` |
| `TableOperationTools.cs` | `add_table_to_model`, `refresh_table` |

---

## Data Flow

### Typical Request Flow

```text
┌───────────┐    MCP/JSON-RPC    ┌──────────────┐
│ AI Client │ ◄──────────────────► MCP Server   │
└───────────┘     (stdio)        │  (Program.cs)│
                                 └──────┬───────┘
                                        │
                                        ▼
                                 ┌──────────────┐
                                 │   MCP Tool   │
                                 │  (Tools/*.cs)│
                                 └──────┬───────┘
                                        │ DI
                                        ▼
                                 ┌──────────────┐
                                 │ Core Service │
                                 │ (Core/*.cs)  │
                                 └──────┬───────┘
                                        │
                                        ▼
                                 ┌──────────────┐
                                 │ Infra Service│
                                 │(Infra/*.cs)  │
                                 └──────┬───────┘
                                        │ STA Dispatch
                                        ▼
                                 ┌──────────────┐
                                 │ PowerPivot   │
                                 │ Connection   │
                                 └──────┬───────┘
                                        │ COM
                                        ▼
                                 ┌──────────────┐
                                 │    Excel     │
                                 │ Power Pivot  │
                                 └──────────────┘
```

### Threading Model

```text
┌─────────────────────────────────────────────────────────┐
│                    MCP Thread Pool                       │
│  (Handles stdio I/O, JSON-RPC, async continuations)     │
└────────────────────────┬────────────────────────────────┘
                         │ Task.Run / await
                         ▼
┌─────────────────────────────────────────────────────────┐
│                   ExcelStaService                        │
│  (Dedicated STA thread with Windows Forms message pump)  │
│  - Executes all COM operations                           │
│  - Manages COM object lifetimes                          │
└────────────────────────┬────────────────────────────────┘
                         │ COM Interop
                         ▼
┌─────────────────────────────────────────────────────────┐
│                    Excel.Application                     │
│  (Single-threaded apartment COM object)                  │
└─────────────────────────────────────────────────────────┘
```

---

## Configuration Management

Configuration is loaded from environment variables with sensible defaults:

```csharp
McpConfiguration
├── ServerName           // MCP_SERVER_NAME
├── ServerVersion        // MCP_SERVER_VERSION
├── MaxQueryRows         // MCP_MAX_QUERY_ROWS (1000)
├── QueryTimeoutSeconds  // MCP_QUERY_TIMEOUT_SECONDS (120)
├── ValidationTimeoutSeconds // MCP_VALIDATION_TIMEOUT_SECONDS (10)
├── DmvTimeoutSeconds    // MCP_DMV_TIMEOUT_SECONDS (30)
├── ConnectionRetryCount // MCP_CONNECTION_RETRY_COUNT (3)
├── ConnectionRetryDelayMs // MCP_CONNECTION_RETRY_DELAY_MS (1000)
└── ConnectionTimeoutMs  // MCP_CONNECTION_TIMEOUT_MS (5000)
```

---

## Error Handling Strategy

1. **Tool-Level**: Exception filter wraps all tools, returns `IsError: true` with message
2. **Connection Recovery**: Detects stale connections, prompts reconnection
3. **COM Error Classification**: `IsFatalComError()` detects unrecoverable errors
4. **Polly Retry Policies**: Exponential backoff for transient failures

---

## Resource Management

MCP Resources are provided via `McpResourceProvider.cs`:

| URI | Description |
| ----- | ------------- |
| `model://schema` | Current Power Pivot model schema (JSON) |
| `model://dirty-state` | Unsaved changes indicator |
| `model://logs` | Recent server diagnostic logs |
| `model://instructions` | Guidelines for working with Power Pivot |
| `model://best-practices` | Measure best practices guide |
| `model://dax-guide` | DAX query guide |
| `model://workflows` | Common workflow reference |
