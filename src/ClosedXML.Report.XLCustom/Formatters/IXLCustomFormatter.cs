namespace ClosedXML.Report.XLCustom.Formatters;

/// <summary>
/// Interface for custom formatters that can be registered with <see cref="XLCustomTemplate"/>
/// </summary>
public interface IXLCustomFormatter
{
    /// <summary>
    /// Formats the specified value according to the format string and parameters
    /// </summary>
    /// <param name="value">The value to format</param>
    /// <param name="parameters">Optional parameters that can be used for formatting</param>
    /// <returns>The formatted value</returns>
    object Format(object value, string[] parameters);
}