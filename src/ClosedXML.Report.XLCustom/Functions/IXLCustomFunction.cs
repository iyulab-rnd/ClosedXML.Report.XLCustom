using ClosedXML.Excel;

namespace ClosedXML.Report.XLCustom.Functions;

/// <summary>
/// Interface for custom functions that can be registered with <see cref="XLCustomTemplate"/>
/// </summary>
public interface IXLCustomFunction
{
    /// <summary>
    /// Processes a cell with a specified value and parameters
    /// </summary>
    /// <param name="cell">The Excel cell to be processed</param>
    /// <param name="value">The value to be processed</param>
    /// <param name="parameters">Optional parameters that can be used for processing</param>
    void Process(IXLCell cell, object value, string[] parameters);
}