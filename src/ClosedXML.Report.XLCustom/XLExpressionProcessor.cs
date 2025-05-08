namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Processes custom expressions and converts them to compatible tags
/// </summary>
public class XLExpressionProcessor
{
    private readonly FunctionRegistry _functionRegistry;
    private readonly GlobalVariableRegistry _globalVariables;

    // Format expression: {{variable:format}} 
    private static readonly Regex FormatExpressionRegex =
        new Regex(@"\{\{\s*([^{}:]+)\s*:\s*([^{}]+)\s*\}\}", RegexOptions.Compiled);

    // Function expression: {{variable|function}} or {{variable|function(parameters)}}
    private static readonly Regex FunctionExpressionRegex =
        new Regex(@"\{\{\s*([^{}|]+)\s*\|\s*([^{}(]+)\s*(?:\(\s*([^{}]*)\s*\))?\s*\}\}", RegexOptions.Compiled);

    // Simple variable expression: {{variable}} - 이제 처리하지 않고 ClosedXML에서 처리하도록 함
    // 하지만 글로벌 변수가 있는지 확인하기 위해 패턴 유지
    private static readonly Regex SimpleVariableRegex =
        new Regex(@"\{\{\s*([^{}|:]+)\s*\}\}", RegexOptions.Compiled);

    public XLExpressionProcessor(FunctionRegistry functionRegistry, GlobalVariableRegistry globalVariables)
    {
        _functionRegistry = functionRegistry ?? throw new ArgumentNullException(nameof(functionRegistry));
        _globalVariables = globalVariables ?? throw new ArgumentNullException(nameof(globalVariables));
    }

    public string ProcessExpression(string value, IXLCell cell)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        Log.Debug($"Processing expression: {value}");

        // 간소화된 접근 방식 - format과 function만 변환
        // 변수 표현식은 그대로 두고 ClosedXML.Report가 처리하도록 함
        return ProcessFormatAndFunctionExpressions(value);
    }

    private string ProcessFormatAndFunctionExpressions(string value)
    {
        bool isModified = false;

        // Format 표현식 처리: {{Value:format}}
        string result = FormatExpressionRegex.Replace(value, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var formatName = match.Groups[2].Value.Trim();
            Log.Debug($"Format expression found: {variableName}:{formatName}");

            // 변수가 실제로 사용 가능한지 확인
            if (!IsValidVariable(variableName))
            {
                Log.Debug($"Variable not found: {variableName}, keeping original expression");
                return match.Value;
            }

            // format 태그 생성
            isModified = true;
            return $"<<format name=\"{variableName}\" format=\"{formatName}\">>";
        });

        if (isModified)
        {
            value = result;
            Log.Debug($"After format processing: {value}");
            isModified = false;
        }

        // Function 표현식 처리: {{Value|function}} 또는 {{Value|function(params)}}
        result = FunctionExpressionRegex.Replace(value, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var functionName = match.Groups[2].Value.Trim();
            var paramString = match.Groups.Count > 3 ? match.Groups[3].Value : "";

            Log.Debug($"Function expression found: {variableName}|{functionName}({paramString})");

            // 변수와 함수가 모두 사용 가능한지 확인
            if (!IsValidVariable(variableName))
            {
                Log.Debug($"Variable not found: {variableName}, keeping original expression");
                return match.Value;
            }

            if (!_functionRegistry.IsRegistered(functionName))
            {
                Log.Debug($"Function not registered: {functionName}, keeping original expression");
                return match.Value;
            }

            // 태그 매개변수 생성
            var tagParams = new StringBuilder();

            // 함수 태그 생성
            tagParams.Append($"<<customfunction name=\"{variableName}\" function=\"{functionName}\"");

            if (!string.IsNullOrEmpty(paramString))
            {
                var parameters = ParseParameters(paramString);
                tagParams.Append($" parameters=\"{EscapeParameter(string.Join(",", parameters))}\"");
            }

            tagParams.Append(">>");

            Log.Debug($"Created function tag for: {functionName}");
            isModified = true;

            return tagParams.ToString();
        });

        if (isModified)
        {
            value = result;
            Log.Debug($"After function processing: {value}");
        }

        return value;
    }

    // 변수가 템플릿에서 사용 가능한지 확인
    private bool IsValidVariable(string variableName)
    {
        // 글로벌 변수인지 확인 (필요시 추가 검증 로직 구현)
        if (_globalVariables.IsRegistered(variableName))
        {
            return true;
        }

        // 점으로 구분된 변수 이름인 경우 (아이템 속성 또는 중첩 객체)
        if (variableName.Contains('.'))
        {
            // item.Property 같은 패턴 검증
            string[] parts = variableName.Split('.');
            if (parts.Length > 1 && parts[0].Equals("item", StringComparison.OrdinalIgnoreCase))
            {
                // item.xxx 형식은 유효한 것으로 간주
                return true;
            }
        }

        // 이 시점에서는 모델 변수를 확인할 수 없으므로
        // 글로벌 변수나 item 속성이 아니면 가능하다고 가정하고 계속 진행
        return true;
    }

    /// <summary>
    /// 간소화된 변수 표현식({{Variable}})을 찾아서 기록
    /// </summary>
    public HashSet<string> GetSimpleVariableNames(string value)
    {
        var result = new HashSet<string>();

        if (string.IsNullOrEmpty(value))
            return result;

        // 간단한 변수 표현식 검색
        var matches = SimpleVariableRegex.Matches(value);
        foreach (Match match in matches)
        {
            if (match.Groups.Count > 1)
            {
                var variableName = match.Groups[1].Value.Trim();
                result.Add(variableName);
            }
        }

        return result;
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
        bool inQuote = false;
        char lastChar = '\0';

        for (int i = 0; i < paramString.Length; i++)
        {
            char c = paramString[i];

            // Handle quoted strings
            if (c == '\'' && lastChar != '\\')
            {
                inQuote = !inQuote;
                currentParam.Append(c);
            }
            // Handle parentheses (only count them if not in a quote)
            else if (c == '(' && !inQuote)
            {
                parenLevel++;
                currentParam.Append(c);
            }
            else if (c == ')' && !inQuote)
            {
                parenLevel--;
                currentParam.Append(c);
            }
            // Handle parameter separator (only if not in quotes or parentheses)
            else if (c == ',' && parenLevel == 0 && !inQuote)
            {
                // Add parameter when a separator is encountered
                parameters.Add(currentParam.ToString().Trim());
                currentParam.Clear();
            }
            else
            {
                currentParam.Append(c);
            }

            lastChar = c;
        }

        // Add final parameter
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
        // Wrap parameter in single quotes if it contains commas or parentheses
        if (param.Contains(',') || param.Contains('(') || param.Contains(')'))
        {
            // Escape existing single quotes
            param = param.Replace("'", "''");
            return $"'{param}'";
        }

        return param;
    }
}