namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Handler for formatting values in expressions with format syntax
    /// </summary>
    /// <param name="value">The value to format</param>
    /// <param name="parameters">Optional parameters for the formatter</param>
    /// <returns>The formatted value</returns>
    public delegate object FormatHandler(object value, string[] parameters);

    /// <summary>
    /// Handler for applying functions to cells in expressions with function syntax
    /// </summary>
    /// <param name="cell">The target cell</param>
    /// <param name="value">The value to process</param>
    /// <param name="parameters">Optional parameters for the function</param>
    public delegate void FunctionHandler(IXLCell cell, object value, string[] parameters);

    /// <summary>
    /// Resolver for undefined variables
    /// </summary>
    /// <param name="expression">The variable expression to resolve</param>
    /// <returns>The resolved value or null if not resolved</returns>
    public delegate object GlobalVariableResolver(string expression);
}