namespace ClosedXML.Report.XLCustom;

internal static class TemplateErrorsExtension
{
    /// <summary>
    /// Adds a simple error message to the errors collection
    /// </summary>
    public static void Add(this TemplateErrors errors, string message)
    {
        try
        {
            // Create a dummy workbook and range
            var tempWorkbook = new XLWorkbook();
            var ws = tempWorkbook.AddWorksheet("Temp");
            var range = ws.Range("A1:A1");

            // Add error with dummy range
            errors.Add(new TemplateError(message, range));
        }
        catch
        {
            // Log to debug as fallback
            Debug.WriteLine($"Template Error: {message}");
        }
    }
}
