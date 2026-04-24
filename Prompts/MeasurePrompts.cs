using System.ComponentModel;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace ExcelPowerPivotMcp.Prompts;

/// <summary>
/// MCP Prompts for common Power Pivot workflows.
/// 
/// Workflow prompts (prefixed with 'workflow_') prescribe multi-step tool sequences.
/// Pure prompts (like 'describe_model') provide context without prescribing actions.
/// </summary>
[McpServerPromptType]
public class MeasurePrompts
{
    #region Pure Prompts (Context only, no tool prescriptions)

    [McpServerPrompt(Name = "describe_model")]
    [Description("Describe the connected Power Pivot data model in plain English. Non-prescriptive - uses context only.")]
    public static IEnumerable<PromptMessage> DescribeModel()
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = """
                        Describe the connected Power Pivot data model in plain English.
                        
                        Provide:
                        - A high-level summary of what this model represents
                        - The main fact tables and their purpose
                        - The dimension tables and their attributes
                        - Key measures and what business questions they answer
                        - The relationship structure (star schema, snowflake, etc.)
                        
                        Be concise but comprehensive.
                        """
                }
            }
        };
    }

    #endregion

    #region Workflow Prompts (Suggested approaches, not strict sequences)

    [McpServerPrompt(Name = "workflow_create_measure")]
    [Description("Create a DAX measure from a natural language description. Provides context and best practices.")]
    public static IEnumerable<PromptMessage> CreateMeasureFromDescription(
        [Description("What the measure should calculate, e.g. 'total sales by category'")] string description,
        [Description("Target table name where the measure will be created")] string tableName)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Create a measure in table '{tableName}' that: {description}

                        CONTEXT:
                        - You'll need the model schema to write correct DAX
                        - The measure will need to be saved after creation
                        
                        DAX BEST PRACTICES:
                        - Use DIVIDE(numerator, denominator, 0) instead of / for safe division
                        - Use VAR for intermediate calculations for readability
                        - Include a description parameter explaining what the measure does
                        
                        See model://best-practices resource for more DAX patterns.
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_analyze_model")]
    [Description("Analyze the current Power Pivot data model structure comprehensively.")]
    public static IEnumerable<PromptMessage> AnalyzeModel()
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = """
                        Analyze this Power Pivot data model comprehensively.
                        
                        PROVIDE:
                        - Summary of the data model (fact tables, dimensions)
                        - List of existing measures grouped by purpose
                        - Relationship structure (star schema, snowflake, etc.)
                        - Any potential issues (missing relationships, orphan tables)
                        - Suggestions for additional measures
                        
                        Use get_model_summary with detail_level='full' for the most complete picture.
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_test_dax")]
    [Description("Test a DAX expression before creating a permanent measure.")]
    public static IEnumerable<PromptMessage> TestDaxExpression(
        [Description("The DAX expression to test")] string expression)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Test this DAX expression before creating a measure:

                        ```dax
                        {expression}
                        ```

                        APPROACH:
                        Use run_dax with a DEFINE MEASURE block to test without persisting:
                        ```
                        DEFINE MEASURE 'TableName'[Test] = {expression}
                        EVALUATE ROW("Result", [Test])
                        ```
                        
                        Report if the expression is valid and what value it returns.
                        If there's an error, explain what's wrong and suggest a fix.
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_fix_measure")]
    [Description("Debug and fix a measure that isn't working correctly.")]
    public static IEnumerable<PromptMessage> FixMeasure(
        [Description("Name of the measure to fix")] string measureName,
        [Description("What's wrong with it")] string problem)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Fix the measure '{measureName}'. Problem: {problem}

                        DEBUGGING TIPS:
                        - Check the current expression with list_measures
                        - Test fixes with run_dax before updating
                        - Remember to save after updating
                        
                        COMMON ISSUES:
                        - Division by zero: Use DIVIDE(x, y, 0)
                        - Blank handling: Use IF(ISBLANK(...), ...)
                        - Filter context: May need CALCULATE or REMOVEFILTERS
                        - Case-sensitive column/table names
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_add_relationship")]
    [Description("Create a relationship between two tables in the data model.")]
    public static IEnumerable<PromptMessage> AddRelationship(
        [Description("The foreign key (many-side) table")] string foreignTable,
        [Description("The primary key (one-side) table")] string primaryTable)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Create a relationship from '{foreignTable}' to '{primaryTable}'.

                        RELATIONSHIP RULES:
                        - Foreign key (many-side): '{foreignTable}' - can have duplicate values
                        - Primary key (one-side): '{primaryTable}' - must have unique values
                        - Columns must have matching data types
                        
                        VALIDATION:
                        - Verify primary key column has unique values (analyze_column can help)
                        - Check columns have compatible data types
                        - Remember to save after creating

                        See model://best-practices for relationship design patterns.
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_add_table")]
    [Description("Add an Excel table to the Power Pivot data model.")]
    public static IEnumerable<PromptMessage> AddTableToModel(
        [Description("Name of the Excel table to add")] string tableName)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Add the Excel table '{tableName}' to the Power Pivot data model.

                        APPROACH:
                        1. Use list_excel_tables to see available tables and their model status
                        2. Use add_excel_table_to_model with confirm=true to add the table
                        3. Use save_workbook to persist the change

                        ALTERNATIVE (with transformations):
                        If transformations are needed, use add_power_query_table_to_model with M-code.
                        
                        After adding, verify with list_tables that the table is in the model.
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_optimize_performance")]
    [Description("Analyze and optimize Power Pivot model performance and memory usage.")]
    public static IEnumerable<PromptMessage> OptimizePerformance()
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = """
                        Analyze this Power Pivot model's performance and suggest optimizations.

                        DIAGNOSTIC STEPS:
                        1. Use analyze_storage_statistics with level='tables' for overview
                        2. Identify largest tables (by memory) and check bytesPerRow
                        3. Use analyze_vertipaq_compression on problematic tables
                        4. Look for high-cardinality text columns and poor compression ratios

                        COMMON OPTIMIZATIONS:
                        - Remove unused columns (check with list_dependencies first)
                        - Convert text IDs to integers
                        - Remove time from date columns if not needed
                        - Split repeated data into dimension tables

                        See model://best-practices resource for detailed optimization patterns.
                        TARGET: bytesPerRow < 50, compression ratio < 0.5
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_refresh_data")]
    [Description("Refresh data model tables from their source connections.")]
    public static IEnumerable<PromptMessage> RefreshData(
        [Description("Optional: specific table to refresh (leave empty for full refresh)")] string? tableName = null)
    {
        var tableSpecific = !string.IsNullOrEmpty(tableName);
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = tableSpecific
                        ? $"""
                            Refresh the table '{tableName}' from its data source.

                            STEPS:
                            1. Use refresh_table with table_name='{tableName}'
                            2. Verify with run_dax to check row counts
                            3. Remember to save_workbook after refresh
                            """
                        : """
                            Refresh all tables in the data model from their sources.

                            STEPS:
                            1. Use refresh_model to refresh all tables
                            2. Verify with get_model_summary to check row counts
                            3. Remember to save_workbook after refresh

                            For selective refresh, use refresh_table with specific table names.
                            Check list_power_queries to understand source connections.
                            """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_delete_measure_safely")]
    [Description("Safely delete a measure after checking for dependencies.")]
    public static IEnumerable<PromptMessage> DeleteMeasureSafely(
        [Description("Name of the measure to delete")] string measureName)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Safely delete the measure '{measureName}' after checking dependencies.

                        SAFETY STEPS:
                        1. Use list_dependencies with object_name='{measureName}' to find dependents
                        2. If other measures depend on it, either:
                           - Update those measures first, OR
                           - Confirm deletion will break them (warn user)
                        3. Use delete_measure with confirm=true to delete
                        4. Use save_workbook to persist

                        WARNING: Deleting a measure that other measures reference will break them!
                        """
                }
            }
        };
    }

    [McpServerPrompt(Name = "workflow_validate_and_create_measure")]
    [Description("Validate a DAX expression first, then create the measure.")]
    public static IEnumerable<PromptMessage> ValidateAndCreateMeasure(
        [Description("DAX expression for the measure")] string expression,
        [Description("Target table for the measure")] string tableName,
        [Description("Name for the measure")] string measureName)
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = $"""
                        Create measure '{measureName}' in table '{tableName}' after validating.

                        Expression: {expression}

                        VALIDATION WORKFLOW:
                        1. Use validate_dax to check syntax of: {expression}
                        2. Test with run_dax using DEFINE MEASURE pattern:
                           DEFINE MEASURE '{tableName}'[{measureName}Test] = {expression}
                           EVALUATE ROW("Result", [{measureName}Test])
                        3. If valid, use create_measure to create the permanent measure
                        4. Use save_workbook to persist

                        DAX BEST PRACTICES:
                        - Use DIVIDE(x, y, 0) instead of x/y for safe division
                        - Use VAR for intermediate calculations
                        - Handle blanks with IF(ISBLANK(...), ...)
                        """
                }
            }
        };
    }

    #endregion

    #region Error Recovery Prompts

    [McpServerPrompt(Name = "recover_connection")]
    [Description("Recover from a lost Excel connection. Use this when you get 'Connection lost' errors.")]
    public static IEnumerable<PromptMessage> RecoverConnection()
    {
        return new[]
        {
            new PromptMessage
            {
                Role = Role.User,
                Content = new TextContentBlock
                {
                    Text = """
                        The connection to Excel was lost. Help me recover.

                        RECOVERY STEPS:
                        1. Use discover_workbooks to find available workbooks
                        2. Use connect_workbook to reconnect
                        3. Verify connection with get_connection_status
                        4. Check model://dirty-state to see if there were unsaved changes

                        COMMON CAUSES:
                        - Excel was closed unexpectedly
                        - The workbook was closed
                        - Excel crashed or became unresponsive
                        - A modal dialog blocked the operation

                        After reconnecting, check if any work was lost.
                        """
                }
            }
        };
    }

    #endregion
}

