namespace ClosedXML.Report.XLCustom.Internals;

/// <summary>
/// Extensions for TemplateErrors handling
/// </summary>
internal static class TemplateErrorExtensions
{
    /// <summary>
    /// Adds a template error with message and optional sheet name
    /// </summary>
    public static void Add(this TemplateErrors errors, string message, string sheetName = null)
    {
        // Create a dummy range using the first available worksheet
        IXLRange dummyRange = null;

        try
        {
            // Create a temporary workbook if needed
            var tempWorkbook = new XLWorkbook();
            var tempWorksheet = tempWorkbook.AddWorksheet("Temp");
            dummyRange = tempWorksheet.Range("A1:A1");

            // Add the error with the dummy range
            errors.Add(new TemplateError(message, dummyRange));
        }
        catch
        {
            // Last resort - if we can't create a range, log to console 
            // but don't throw an exception as this would disrupt template processing
            Debug.WriteLine($"Template Error: {message} in sheet: {sheetName}");
        }
    }
}