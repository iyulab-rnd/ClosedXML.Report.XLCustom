namespace ClosedXML.Report.XLCustom.Internals
{
    /// <summary>
    /// Custom formula evaluator that extends the base FormulaEvaluator with format and function capabilities
    /// </summary>
    internal class CustomFormulaEvaluator
    {
        private readonly FormulaEvaluator _baseEvaluator;
        private readonly Dictionary<string, FormatHandler> _formatters;
        private readonly Dictionary<string, FunctionHandler> _functions;
        private GlobalVariableResolver _globalResolver;

        public CustomFormulaEvaluator(FormulaEvaluator baseEvaluator,
                                    Dictionary<string, FormatHandler> formatters,
                                    Dictionary<string, FunctionHandler> functions)
        {
            _baseEvaluator = baseEvaluator;
            _formatters = formatters;
            _functions = functions;
        }

        // 글로벌 리졸버 설정
        public void SetGlobalResolver(GlobalVariableResolver resolver)
        {
            _globalResolver = resolver;
        }

        // 표현식 평가
        public object Evaluate(string formula, IXLCell cell, params Parameter[] pars)
        {
            // 표현식 파싱
            var expression = ExpressionParser.Parse(formula);

            // 1. 기본 값 평가
            object value = _baseEvaluator.Evaluate(expression.Value, pars);

            // 2. 글로벌 리졸버 시도
            if (value == null && _globalResolver != null)
            {
                var variableName = ExtractVariableName(expression.Value);
                if (!string.IsNullOrEmpty(variableName))
                {
                    value = _globalResolver(variableName);
                }
            }

            // 3. 포맷 적용
            if (expression.HasFormat && _formatters.TryGetValue(expression.Format.ToLowerInvariant(), out var formatter))
            {
                value = formatter(value, expression.FormatParameters);
            }

            // 4. 함수 적용
            if (expression.HasFunction && _functions.TryGetValue(expression.Function.ToLowerInvariant(), out var function))
            {
                function(cell, value, expression.FunctionParameters);
            }

            return value;
        }

        // 변수 이름 추출
        private string ExtractVariableName(string expression)
        {
            if (string.IsNullOrEmpty(expression)) return null;

            if (expression.StartsWith("{{") && expression.EndsWith("}}"))
            {
                var content = expression.Substring(2, expression.Length - 4).Trim();

                // 간단한 변수 이름인지 확인
                if (!content.Contains(".") && !content.Contains("(") &&
                    !content.Contains("+") && !content.Contains("-") &&
                    !content.Contains("*") && !content.Contains("/"))
                {
                    return content;
                }
            }

            return null;
        }
    }
}