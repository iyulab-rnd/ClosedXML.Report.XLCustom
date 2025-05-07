using System.Text.RegularExpressions;

namespace ClosedXML.Report.XLCustom.Parsing
{
    /// <summary>
    /// Parses custom expression syntax including formats and functions
    /// </summary>
    internal static class ExpressionParser
    {
        // 정규식 컴파일하여 성능 향상
        private static readonly Regex FormatRegex = new(@"{{(.+?):(.+?)}}", RegexOptions.Compiled);
        private static readonly Regex FunctionRegex = new(@"{{(.+?)\|(.+?)}}", RegexOptions.Compiled);
        private static readonly Regex ParameterRegex = new(@"^([^(]+)(?:\((.+)\))?$", RegexOptions.Compiled);

        // 표현식 파싱
        public static Expression Parse(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return new Expression { OriginalExpression = input, Value = input };
            }

            // 포맷 표현식 확인
            var formatMatch = FormatRegex.Match(input);
            if (formatMatch.Success)
            {
                var expr = new Expression
                {
                    OriginalExpression = input,
                    Value = formatMatch.Groups[1].Value,
                    HasFormat = true
                };

                var formatExpr = formatMatch.Groups[2].Value;
                ParseNameAndParameters(formatExpr, out string formatName, out string[] parameters);

                expr.Format = formatName;
                expr.FormatParameters = parameters;

                return expr;
            }

            // 함수 표현식 확인
            var functionMatch = FunctionRegex.Match(input);
            if (functionMatch.Success)
            {
                var expr = new Expression
                {
                    OriginalExpression = input,
                    Value = functionMatch.Groups[1].Value,
                    HasFunction = true
                };

                var functionExpr = functionMatch.Groups[2].Value;
                ParseNameAndParameters(functionExpr, out string functionName, out string[] parameters);

                expr.Function = functionName;
                expr.FunctionParameters = parameters;

                return expr;
            }

            // 일반 표현식
            return new Expression
            {
                OriginalExpression = input,
                Value = input
            };
        }

        // 이름과 매개변수 파싱
        private static void ParseNameAndParameters(string expression, out string name, out string[] parameters)
        {
            var match = ParameterRegex.Match(expression);

            if (match.Success)
            {
                name = match.Groups[1].Value.Trim();

                if (match.Groups.Count > 2 && !string.IsNullOrEmpty(match.Groups[2].Value))
                {
                    var paramString = match.Groups[2].Value;
                    parameters = SplitParameters(paramString);
                }
                else
                {
                    parameters = Array.Empty<string>();
                }
            }
            else
            {
                name = expression.Trim();
                parameters = Array.Empty<string>();
            }
        }

        // 매개변수 분할 (쉼표 구분)
        private static string[] SplitParameters(string paramString)
        {
            var result = new List<string>();
            var builder = new StringBuilder();
            var inQuote = false;

            for (int i = 0; i < paramString.Length; i++)
            {
                char c = paramString[i];

                if (c == '"')
                {
                    inQuote = !inQuote;
                    builder.Append(c);
                }
                else if (c == ',' && !inQuote)
                {
                    result.Add(builder.ToString().Trim());
                    builder.Clear();
                }
                else
                {
                    builder.Append(c);
                }
            }

            if (builder.Length > 0)
            {
                result.Add(builder.ToString().Trim());
            }

            return result.ToArray();
        }
    }

    // 표현식 클래스
    internal class Expression
    {
        public string OriginalExpression { get; set; }
        public string Value { get; set; }

        public bool HasFormat { get; set; }
        public string Format { get; set; }
        public string[] FormatParameters { get; set; } = Array.Empty<string>();

        public bool HasFunction { get; set; }
        public string Function { get; set; }
        public string[] FunctionParameters { get; set; } = Array.Empty<string>();
    }
}