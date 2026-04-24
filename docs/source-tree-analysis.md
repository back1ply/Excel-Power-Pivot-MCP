# Excel Power Pivot MCP - Source Tree Analysis

## Directory Structure

```text
Excel-Power-Pivot-MCP/                     # Project root
│
├── Program.cs                             # 🚀 Entry point - DI setup, MCP server config
├── McpConfiguration.cs                    # ⚙️ Configuration from env vars
├── McpResourceProvider.cs                 # 📦 MCP resource providers
├── ExcelPowerPivotMcp.csproj             # 📋 Project file (.NET 8)
├── ExcelPowerPivotMcp.sln                # 📋 Solution file
│
├── Core/                                  # 🔷 CORE LAYER - Business logic
│   └── Services/                          # Service interfaces & implementations
│       ├── Interfaces.cs                  # IMeasureService, IRelationshipService, etc.
│       ├── MeasureService.cs              # Measure CRUD implementation
│       ├── RelationshipService.cs         # Relationship management
│       ├── TableService.cs                # Table operations
│       ├── ModelMetadataService.cs        # Metadata queries
│       └── DataProfileService.cs          # Data profiling
│
├── Infrastructure/                        # 🔶 INFRASTRUCTURE LAYER - External integrations
│   └── Services/
│       ├── IInfrastructureServices.cs     # Interface definitions
│       ├── ExcelInteropService.cs         # Excel Application interactions
│       ├── ExcelStaService.cs             # STA thread message pump
│       ├── PowerPivotComService.cs        # Power Pivot COM model (28KB - largest file)
│       ├── PowerPivotConnectionManager.cs # Connection singleton management
│       ├── AdoDmvService.cs               # DMV metadata queries
│       ├── AdoDaxService.cs               # DAX query execution
│       └── InMemoryLogger.cs              # Log capture for MCP resource
│
├── PowerPivot/                            # 🔌 POWERPIVOT LAYER - Connection management
│   ├── PowerPivotConnection.cs            # Main connection class (partial)
│   ├── PowerPivotConnection.Discovery.cs  # Workbook discovery logic
│   ├── ComObjectManager.cs                # COM object lifecycle
│   ├── DictionaryExtensions.cs            # Helper extensions
│   ├── ExcelComHelpers.cs                 # Excel COM utilities
│   └── ExceptionHelpers.cs                # Exception handling
│
├── Tools/                                 # 🛠️ MCP TOOLS - API endpoints
│   ├── ConnectionTools.cs                 # discover_workbooks, connect_workbook, save, refresh
│   ├── DaxQueryTools.cs                   # run_dax, analyze_column
│   ├── MeasureCrudTools.cs                # create/update/delete measure
│   ├── ModelMetadataTools.cs              # list_tables, list_columns, get_model_summary
│   ├── RelationshipCrudTools.cs           # create/delete relationship, set_active
│   └── TableOperationTools.cs             # add_table_to_model, refresh_table
│
├── Prompts/                               # 💬 MCP PROMPTS
│   └── MeasurePrompts.cs                  # Guided measure creation workflows
│
├── Resources/                             # 📚 EMBEDDED DOCUMENTATION
│   ├── common_workflows.md                # Common workflow reference
│   ├── dax_query_excel_guide.md           # DAX query guide for Excel
│   ├── excel_powerpivot_instructions.md   # Power Pivot guidelines
│   └── powerpivot_measure_best_practices.md # Measure best practices
│
├── Common/                                # 🔧 SHARED UTILITIES
│   ├── EmbeddedResources.cs               # Resource loading helper
│   ├── DataStructures/                    # DTOs and data models
│   │   ├── MeasureCreate.cs               # Create measure request
│   │   ├── MeasureUpdate.cs               # Update measure request
│   │   ├── RelationshipCreate.cs          # Create relationship request
│   │   ├── TableAddToModel.cs             # Add table request
│   │   └── Metadata/                      # Metadata response models
│   └── Utils/
│       └── DaxHelpers.cs                  # DAX-related utilities
│
├── Tests/                                 # 🧪 UNIT TESTS
│   ├── ExcelPowerPivotMcp.Tests.csproj    # Test project
│   └── DaxHelpersTests.cs                 # DAX helper tests
│
├── docs/                                  # 📖 PROJECT DOCUMENTATION
│   ├── index.md                           # Documentation index (this file)
│   ├── project-overview.md                # Project summary
│   ├── architecture.md                    # Architecture details
│   └── dmv_coverage.md                    # DMV query coverage
│
└── [Build artifacts]
    ├── bin/                               # Build output
    ├── obj/                               # Build intermediates
    ├── build.ps1                          # Build script
    └── release.ps1                        # Release script
```

---

## Critical Directories Explained

### `/Core/Services`

**Purpose:** Business logic layer with domain services.

All services are:

- Interface-based for testability
- Async-first to avoid blocking
- Injected via DI container

### `/Infrastructure/Services`

**Purpose:** External integrations and platform-specific code.

Key files:

- `ExcelStaService.cs` - Dedicated STA thread with Windows Forms message pump for COM
- `PowerPivotComService.cs` - Largest file (28KB) - all COM operations for Power Pivot
- `AdoDmvService.cs` - ADOMD.NET queries for metadata
- `AdoDaxService.cs` - ADOMD.NET queries for data

### `/PowerPivot`

**Purpose:** Excel connection and COM object management.

- `PowerPivotConnection.cs` is a partial class split for organization
- Handles connection lifecycle, validation, and recovery
- Manages ADOMD.NET connection to Power Pivot's internal SSAS instance

### `/Tools`

**Purpose:** MCP tool implementations auto-discovered via reflection.

Each tool file contains multiple `[McpServerTool]` attributed methods that become
available to AI clients through the MCP protocol.

### `/Resources`

**Purpose:** Embedded markdown documentation.

These files are:

- Embedded in the assembly at compile time
- Served as MCP resources via `model://` URI scheme
- Available to AI clients for context about Power Pivot

---

## File Size Analysis

| File | Size | Notes |
| ---- | ---- | ----- |
| `PowerPivotComService.cs` | ~28KB | Largest - all COM operations |
| `PowerPivotConnection.cs` | ~15.5KB | Connection management |
| `excel_powerpivot_instructions.md` | ~15KB | Comprehensive AI instructions |
| `ModelMetadataTools.cs` | ~15KB | Many metadata query tools |
| `MeasureCrudTools.cs` | ~12.7KB | Measure operations |
| `test_driver.py` | ~63KB | Python test driver (external) |

---

## Entry Points

| Entry Point | File | Purpose |
| ------------- | ------ | --------- |
| **Main** | `Program.cs` | Application startup |
| **Tools** | `Tools/*.cs` | MCP tool handlers |
| **Resources** | `McpResourceProvider.cs` | MCP resource handlers |
| **Prompts** | `Prompts/*.cs` | MCP prompt handlers |
