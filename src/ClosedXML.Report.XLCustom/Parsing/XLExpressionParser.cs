namespace ClosedXML.Report.XLCustom.Parsing;

/// <summary>
/// Parser for enhanced expression syntax
/// </summary>
internal static class XLExpressionParser
{
    private static readonly Regex ExpressionRegex = new Regex(@"\{\{([^{}]+)\}\}", RegexOptions.Compiled);
    private static readonly Regex FormatRegex = new Regex(@"^(.+?):(.+)$", RegexOptions.Compiled);
    private static readonly Regex FunctionRegex = new Regex(@"^(.+?)\|(.+)$", RegexOptions.Compiled);
    private static readonly Regex ParametersRegex = new Regex(@"^(\w+)(?:\(([^)]*)\))?$", RegexOptions.Compiled);

    /// <summary>
    /// Extracts all expressions from a given text
    /// </summary>
    public static IEnumerable<string> ExtractExpressions(string text)
    {
        if (string.IsNullOrEmpty(text))
            yield break;

        foreach (Match match in ExpressionRegex.Matches(text))
        {
            yield return match.Value;
        }
    }

    /// <summary>
    /// Parses an expression into its components
    /// </summary>
    public static ParsedExpression Parse(string expression)
    {
        if (string.IsNullOrEmpty(expression))
            throw new ArgumentException("Expression cannot be null or empty", nameof(expression));

        // Extract content between {{ and }}
        var match = ExpressionRegex.Match(expression);
        if (!match.Success)
            throw new ArgumentException($"Invalid expression format: {expression}", nameof(expression));

        var content = match.Groups[1].Value.Trim();

        // Check if it's a function expression (contains | )
        var functionMatch = FunctionRegex.Match(content);
        if (functionMatch.Success)
        {
            var variable = functionMatch.Groups[1].Value.Trim();
            var functionPart = functionMatch.Groups[2].Value.Trim();

            var paramsMatch = ParametersRegex.Match(functionPart);
            if (paramsMatch.Success)
            {
                var functionName = paramsMatch.Groups[1].Value;
                var parameters = paramsMatch.Groups[2].Success
                    ? ParseParameters(paramsMatch.Groups[2].Value)
                    : Array.Empty<string>();

                return new ParsedExpression(
                    XLExpressionType.Function,
                    variable,
                    functionName,
                    parameters,
                    expression);
            }

            // If no parameters pattern, treat the whole thing as the function name
            return new ParsedExpression(
                XLExpressionType.Function,
                variable,
                functionPart,
                Array.Empty<string>(),
                expression);
        }

        // Check if it's a format expression (contains : )
        var formatMatch = FormatRegex.Match(content);
        if (formatMatch.Success)
        {
            var variable = formatMatch.Groups[1].Value.Trim();
            var formatPart = formatMatch.Groups[2].Value.Trim();

            var paramsMatch = ParametersRegex.Match(formatPart);
            if (paramsMatch.Success)
            {
                var formatName = paramsMatch.Groups[1].Value;
                var parameters = paramsMatch.Groups[2].Success
                    ? ParseParameters(paramsMatch.Groups[2].Value)
                    : Array.Empty<string>();

                return new ParsedExpression(
                    XLExpressionType.Format,
                    variable,
                    formatName,
                    parameters,
                    expression);
            }

            // If no parameters pattern, treat the whole thing as the format name
            return new ParsedExpression(
                XLExpressionType.Format,
                variable,
                formatPart,
                Array.Empty<string>(),
                expression);
        }

        // Special case for collections metadata 
        if (content.Contains(".Count"))
        {
            string[] parts = content.Split(new string[] { "." }, StringSplitOptions.None);
            if (parts.Length == 2 && parts[1] == "Count")
            {
                // Convert to variable name
                string countVarName = $"{parts[0]}_Count";

                return new ParsedExpression(
                    XLExpressionType.Standard,
                    countVarName,
                    null,
                    Array.Empty<string>(),
                    expression);
            }
        }

        // Standard variable expression
        return new ParsedExpression(
            XLExpressionType.Standard,
            content,
            null,
            Array.Empty<string>(),
            expression);
    }

    /// <summary>
    /// Parses a parameter string into individual parameters
    /// </summary>
    private static string[] ParseParameters(string parametersString)
    {
        if (string.IsNullOrEmpty(parametersString))
            return Array.Empty<string>();

        // Split by commas, but handle key=value pairs and quoted strings
        var parameters = new List<string>();
        var current = "";
        var inQuotes = false;

        for (int i = 0; i < parametersString.Length; i++)
        {
            char c = parametersString[i];

            if (c == '"' && (i == 0 || parametersString[i - 1] != '\\'))
                inQuotes = !inQuotes;

            if (c == ',' && !inQuotes)
            {
                parameters.Add(current.Trim());
                current = "";
                continue;
            }

            current += c;
        }

        if (!string.IsNullOrEmpty(current))
            parameters.Add(current.Trim());

        return parameters.Select(p => p.Trim()).ToArray();
    }

    /// <summary>
    /// Checks if the text contains an enhanced expression
    /// </summary>
    public static bool IsEnhancedExpression(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        return text.Contains("{{") && (text.Contains(":") || text.Contains("|"));
    }
}

/// <summary>
/// Represents the type of expression
/// </summary>
internal enum XLExpressionType
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

/// <summary>
/// Represents a parsed expression from a template
/// </summary>
internal class ParsedExpression
{
    /// <summary>
    /// Gets the expression type
    /// </summary>
    public XLExpressionType Type { get; }

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
        XLExpressionType type,
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