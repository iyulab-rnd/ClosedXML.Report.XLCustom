namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Delegate for custom format handlers
    /// </summary>
    /// <param name="value">The value to format</param>
    /// <param name="parameters">Optional parameters for the formatter</param>
    /// <returns>The formatted value</returns>
    public delegate object XLFormatHandler(object value, string[] parameters);

    /// <summary>
    /// Delegate for custom function handlers
    /// </summary>
    /// <param name="cell">The cell being processed</param>
    /// <param name="value">The value to process</param>
    /// <param name="parameters">Optional parameters for the function</param>
    public delegate void XLFunctionHandler(IXLCell cell, object value, string[] parameters);
}