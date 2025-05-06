namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Delegate for custom formatters that can be registered with XLCustomTemplate
/// </summary>
public delegate object XLCustomFormatterFunc(object value, string[] parameters);

/// <summary>
/// Delegate for custom functions that can be registered with XLCustomTemplate
/// </summary>
public delegate void XLCustomFunctionFunc(IXLCell cell, object value, string[] parameters);