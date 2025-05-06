namespace ClosedXML.Report.XLCustom.Parsing;

/// <summary>
/// Represents a parsed expression from a template
/// </summary>
internal class ParsedExpression
{
    /// <summary>
    /// Gets the expression type
    /// </summary>
    public ExpressionType Type { get; }

    /// <summary>
    /// Gets the variable part of the expression
    /// </summary>
    public string Variable { get; }

    /// <summary>
    /// Gets the format or function name
    /// </summary>
    public string Operation { get; }

    /// <summary>
    /// Gets the parameters for the format or function
    /// </summary>
    public string[] Parameters { get; }

    /// <summary>
    /// Gets the original expression string
    /// </summary>
    public string OriginalExpression { get; }

    public ParsedExpression(
        ExpressionType type,
        string variable,
        string operation,
        string[] parameters,
        string originalExpression)
    {
        Type = type;
        Variable = variable;
        Operation = operation;
        Parameters = parameters;
        OriginalExpression = originalExpression;
    }
}