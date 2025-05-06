using System.Text.RegularExpressions;

namespace ClosedXML.Report.XLCustom.Parsing;

/// <summary>
/// Parser for enhanced expression syntax
/// </summary>
internal static class ExpressionParser
{
    private static readonly Regex ExpressionRegex = new(@"\{\{([^{}]+)\}\}", RegexOptions.Compiled);
    private static readonly Regex FormatRegex = new(@"^(.+?):(.+)$", RegexOptions.Compiled);
    private static readonly Regex FunctionRegex = new(@"^(.+?)\|(.+)$", RegexOptions.Compiled);
    private static readonly Regex ParametersRegex = new(@"^(\w+)(?:\(([^)]*)\))?$", RegexOptions.Compiled);

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

        Debug.WriteLine($"Parsing expression: {expression}");

        // Extract content between {{ and }}
        var match = ExpressionRegex.Match(expression);
        if (!match.Success)
            throw new ArgumentException($"Invalid expression format: {expression}", nameof(expression));

        var content = match.Groups[1].Value.Trim();
        Debug.WriteLine($"Extracted content: {content}");

        // Check if it's a function expression (contains | )
        var functionMatch = FunctionRegex.Match(content);
        if (functionMatch.Success)
        {
            var variable = functionMatch.Groups[1].Value.Trim();
            var functionPart = functionMatch.Groups[2].Value.Trim();
            Debug.WriteLine($"Found function expression: variable={variable}, function={functionPart}");

            var paramsMatch = ParametersRegex.Match(functionPart);
            if (paramsMatch.Success)
            {
                var functionName = paramsMatch.Groups[1].Value;
                var parameters = paramsMatch.Groups[2].Success
                    ? ParseParameters(paramsMatch.Groups[2].Value)
                    : Array.Empty<string>();

                Debug.WriteLine($"Parsed function: name={functionName}, params={(parameters.Length > 0 ? string.Join(", ", parameters) : "none")}");

                return new ParsedExpression(
                    ExpressionType.Function,
                    variable,
                    functionName,
                    parameters,
                    expression);
            }

            // If no parameters pattern, treat the whole thing as the function name
            Debug.WriteLine($"Parsed function without params: name={functionPart}");
            return new ParsedExpression(
                ExpressionType.Function,
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
            Debug.WriteLine($"Found format expression: variable={variable}, format={formatPart}");

            var paramsMatch = ParametersRegex.Match(formatPart);
            if (paramsMatch.Success)
            {
                var formatName = paramsMatch.Groups[1].Value;
                var parameters = paramsMatch.Groups[2].Success
                    ? ParseParameters(paramsMatch.Groups[2].Value)
                    : Array.Empty<string>();

                Debug.WriteLine($"Parsed format: name={formatName}, params={(parameters.Length > 0 ? string.Join(", ", parameters) : "none")}");

                return new ParsedExpression(
                    ExpressionType.Format,
                    variable,
                    formatName,
                    parameters,
                    expression);
            }

            // If no parameters pattern, treat the whole thing as the format name
            Debug.WriteLine($"Parsed format without params: name={formatPart}");
            return new ParsedExpression(
                ExpressionType.Format,
                variable,
                formatPart,
                Array.Empty<string>(),
                expression);
        }

        // Standard variable expression
        Debug.WriteLine($"Parsed standard variable: {content}");
        return new ParsedExpression(
            ExpressionType.Standard,
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

        Debug.WriteLine($"Parsing parameters: {parametersString}");

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

        var result = parameters.Select(p => p.Trim()).ToArray();
        Debug.WriteLine($"Parsed {result.Length} parameters: {string.Join(", ", result)}");
        return result;
    }

    public static bool IsEnhancedExpression(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        return text.Contains("{{") && (text.Contains(":") || text.Contains("|"));
    }
}