namespace ClosedXML.Report.XLCustom.Parsing;

/// <summary>
/// Represents the type of expression
/// </summary>
internal enum ExpressionType
{
    /// <summary>
    /// Standard variable expression (e.g., {{VariableName}})
    /// </summary>
    Standard,

    /// <summary>
    /// Format expression (e.g., {{Value:format}})
    /// </summary>
    Format,

    /// <summary>
    /// Function expression (e.g., {{Value|function}})
    /// </summary>
    Function
}