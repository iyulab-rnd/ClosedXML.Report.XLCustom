namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Debug and diagnostics functionality for XLCustomTemplate
    /// </summary>
    public partial class XLCustomTemplate
    {
        /// <summary>
        /// Enables detailed diagnostics logging
        /// </summary>
        public static bool EnableDetailedLogging { get; set; } = false;

        /// <summary>
        /// Logs diagnostic information about the template
        /// </summary>
        public void LogDiagnostics()
        {
            Debug.WriteLine("\n==== XLCustomTemplate Diagnostics ====");

            // Log workbook info
            Debug.WriteLine($"Workbook: Worksheets={Workbook.Worksheets.Count}");

            // Log registered variables
            Debug.WriteLine("Registered Variables:");
            foreach (var variable in _variables)
            {
                string valueType = variable.Value?.GetType().Name ?? "null";
                string valueStr = variable.Value?.ToString() ?? "null";
                if (valueStr.Length > 50) valueStr = valueStr.Substring(0, 47) + "...";
                Debug.WriteLine($"  {variable.Key} ({valueType}): {valueStr}");
            }

            // Log formatters
            Debug.WriteLine($"Registered Formatters: {_formatters.Count + _formatDelegates.Count}");
            foreach (var name in GetRegisteredFormatters())
            {
                Debug.WriteLine($"  {name}");
            }

            // Log functions
            Debug.WriteLine($"Registered Functions: {_functions.Count + _functionDelegates.Count}");
            foreach (var name in GetRegisteredFunctions())
            {
                Debug.WriteLine($"  {name}");
            }

            // Log errors
            Debug.WriteLine($"Template Errors: {_errors.Count}");
            foreach (var error in _errors)
            {
                Debug.WriteLine($"  {error.Message} at {error.Range?.RangeAddress}");
            }

            Debug.WriteLine("======================================\n");
        }

        /// <summary>
        /// Diagnostically scan all worksheets for potential template issues
        /// </summary>
        public void DiagnoseTemplateIssues()
        {
            Debug.WriteLine("\n==== Template Diagnostics ====");

            // Scan each worksheet
            foreach (var worksheet in Workbook.Worksheets)
            {
                Debug.WriteLine($"Analyzing worksheet: {worksheet.Name}");

                // Count template expressions
                int standardCount = 0;
                int formatCount = 0;
                int functionCount = 0;

                var templateCells = worksheet.CellsUsed(c => c.GetString().Contains("{{"));

                foreach (var cell in templateCells)
                {
                    string value = cell.GetString();
                    if (value.Contains("{{") && value.Contains("}}"))
                    {
                        if (value.Contains("|"))
                        {
                            functionCount++;
                            Debug.WriteLine($"  Function: {cell.Address} - {value}");
                        }
                        else if (value.Contains(":"))
                        {
                            formatCount++;
                            Debug.WriteLine($"  Format: {cell.Address} - {value}");
                        }
                        else
                        {
                            standardCount++;
                        }
                    }
                }

                Debug.WriteLine($"  Standard expressions: {standardCount}");
                Debug.WriteLine($"  Format expressions: {formatCount}");
                Debug.WriteLine($"  Function expressions: {functionCount}");

                // Scan for ranges
                var rangeTags = worksheet.CellsUsed(c => c.GetString().Contains("<<Range"));
                Debug.WriteLine($"  Range tags: {rangeTags.Count()}");

                foreach (var cell in rangeTags)
                {
                    Debug.WriteLine($"  Range tag: {cell.Address} - {cell.GetString()}");
                }
            }

            Debug.WriteLine("=============================\n");
        }
    }
}