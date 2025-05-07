using ClosedXML.Excel;
using System;
using System.Text.RegularExpressions;
using System.Linq;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Processes custom expressions and converts them to compatible tags
/// </summary>
public class XLExpressionProcessor
{
    private readonly FormatRegistry _formatRegistry;
    private readonly FunctionRegistry _functionRegistry;

    // 포맷 표현식: {{variable:format}} 형태 - 파라미터 없음
    private static readonly Regex FormatExpressionRegex =
        new Regex(@"\{\{([^{}:]+):([^{}]+)\}\}", RegexOptions.Compiled);

    // 함수 표현식: {{variable|function}} 또는 {{variable|function(parameters)}} 형태
    private static readonly Regex FunctionExpressionRegex =
        new Regex(@"\{\{([^{}|]+)\|([^{}(]+)(?:\(([^{}]*)\))?\}\}", RegexOptions.Compiled);

    public XLExpressionProcessor(FormatRegistry formatRegistry, FunctionRegistry functionRegistry)
    {
        _formatRegistry = formatRegistry ?? throw new ArgumentNullException(nameof(formatRegistry));
        _functionRegistry = functionRegistry ?? throw new ArgumentNullException(nameof(functionRegistry));
    }

    public string ProcessExpression(string value, IXLCell cell)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        Log.Debug($"Processing expression: {value}");

        // Process format expressions: {{Value:format}}
        bool isModified = false;
        string result = FormatExpressionRegex.Replace(value, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var formatName = match.Groups[2].Value.Trim();
            Log.Debug($"Format expression found: {variableName}:{formatName}");

            // Standard .NET formats (numeric/date)
            if (IsStandardFormat(formatName))
            {
                Log.Debug($"Standard format: {formatName}");
                isModified = true;
                // 중요: 형식 문자열을 그대로 전달하여 원래 포맷이 적용되도록 함
                return $"<<format name=\"{variableName}\" format=\"{formatName}\">>";
            }

            // Custom formats
            if (_formatRegistry.IsRegistered(formatName))
            {
                Log.Debug($"Custom format: {formatName}");
                isModified = true;
                return $"<<customformat name=\"{variableName}\" format=\"{formatName}\">>";
            }

            // Unregistered format - keep original expression
            Log.Debug($"Unrecognized format: {formatName}, keeping original expression");
            return match.Value;
        });


        if (isModified)
        {
            value = result;
            Log.Debug($"After format processing: {value}");
        }

        // Process function expressions: {{Value|function}} or {{Value|function(params)}}
        isModified = false;
        result = FunctionExpressionRegex.Replace(value, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var functionName = match.Groups[2].Value.Trim();
            var paramString = match.Groups.Count > 3 ? match.Groups[3].Value : "";

            Log.Debug($"Function expression found: {variableName}|{functionName}({paramString})");

            if (_functionRegistry.IsRegistered(functionName))
            {
                // 이름 기반 파라미터로 변경
                var tagParams = new StringBuilder();
                tagParams.Append($"<<CustomFunction name=\"{variableName}\" function=\"{functionName}\"");

                if (!string.IsNullOrEmpty(paramString))
                {
                    var parameters = ParseParameters(paramString);
                    tagParams.Append($" parameters=\"{EscapeParameter(string.Join(",", parameters))}\"");
                }

                tagParams.Append(">>");

                Log.Debug($"Registered function: {functionName}");
                isModified = true;

                return tagParams.ToString();
            }

            // Unregistered function - keep original expression
            Log.Debug($"Unrecognized function: {functionName}, keeping original expression");
            return match.Value;
        });

        if (isModified && result != value)
        {
            Log.Debug($"After function processing: {result}");
            value = result;
        }

        return value;
    }

    /// <summary>
    /// Parses a comma-separated parameter string, properly handling parameters with commas and parentheses
    /// </summary>
    private string[] ParseParameters(string paramString)
    {
        if (string.IsNullOrEmpty(paramString))
            return Array.Empty<string>();

        var parameters = new List<string>();
        var currentParam = new StringBuilder();
        int parenLevel = 0;

        for (int i = 0; i < paramString.Length; i++)
        {
            char c = paramString[i];

            if (c == '(')
            {
                parenLevel++;
                currentParam.Append(c);
            }
            else if (c == ')')
            {
                parenLevel--;
                currentParam.Append(c);
            }
            else if (c == ',' && parenLevel == 0)
            {
                // 매개변수 구분자를 만났을 때 현재 매개변수 추가
                parameters.Add(currentParam.ToString().Trim());
                currentParam.Clear();
            }
            else
            {
                currentParam.Append(c);
            }
        }

        // 마지막 매개변수 추가
        if (currentParam.Length > 0)
        {
            parameters.Add(currentParam.ToString().Trim());
        }

        return parameters.ToArray();
    }

    /// <summary>
    /// Escapes a parameter value for safe inclusion in tag parameters
    /// </summary>
    private string EscapeParameter(string param)
    {
        // 매개변수에 쉼표나 괄호가 포함된 경우 홑따옴표로 묶음
        if (param.Contains(',') || param.Contains('(') || param.Contains(')'))
        {
            // 홑따옴표가 이미 있다면 이스케이프 처리
            param = param.Replace("'", "''");
            return $"'{param}'";
        }

        return param;
    }

    /// <summary>
    /// Checks if a format is a standard .NET format
    /// </summary>
    private bool IsStandardFormat(string format)
    {
        if (string.IsNullOrEmpty(format))
            return false;

        // Common standard formats
        var standardFormats = new[] {
            "C", "D", "E", "F", "G", "N", "P", "R", "X",
            "c", "d", "e", "f", "g", "n", "p", "r", "x",
            // Date and time formats
            "d", "D", "t", "T", "f", "F", "g", "G", "M", "O", "R", "s", "u", "U", "Y"
        };

        // Check if it starts with a standard format
        foreach (var std in standardFormats)
        {
            if (format.Equals(std, StringComparison.OrdinalIgnoreCase) ||
                format.StartsWith(std, StringComparison.OrdinalIgnoreCase) &&
                 format.Length > std.Length &&
                 char.IsDigit(format[std.Length]))
                return true;
        }

        // Check for custom date formats (with year, month, day patterns)
        if (format.Contains("yyyy") || format.Contains("MM") || format.Contains("dd") ||
            format.Contains("HH") || format.Contains("mm") || format.Contains("ss"))
            return true;

        return false;
    }

    private bool IsDateFormat(string formatString)
    {
        if (string.IsNullOrEmpty(formatString))
            return false;

        // 명확한 날짜/시간 포맷 패턴 확인
        string[] datePatterns = new[] { "yyyy", "yy", "MM", "M", "dd", "d",
                                   "HH", "H", "hh", "h", "mm", "m", "ss", "s",
                                   "tt", "t", "fff", "ff", "f" };

        foreach (var pattern in datePatterns)
        {
            if (formatString.Contains(pattern))
                return true;
        }

        return false;
    }
}