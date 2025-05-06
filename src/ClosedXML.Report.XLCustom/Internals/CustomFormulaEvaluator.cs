namespace ClosedXML.Report.XLCustom.Internals
{
    /// <summary>
    /// Enhanced formula evaluator for custom templates
    /// </summary>
    internal class CustomFormulaEvaluator
    {
        private readonly XLCustomTemplate _template;
        private readonly FormulaEvaluator _baseEvaluator;

        public CustomFormulaEvaluator(XLCustomTemplate template, FormulaEvaluator baseEvaluator)
        {
            _template = template ?? throw new ArgumentNullException(nameof(template));
            _baseEvaluator = baseEvaluator ?? throw new ArgumentNullException(nameof(baseEvaluator));
        }

        /// <summary>
        /// Evaluates a formula, including enhanced expressions
        /// </summary>
        public object Evaluate(string formula, IXLCell cell, params Parameter[] parameters)
        {
            if (string.IsNullOrEmpty(formula))
            {
                Debug.WriteLine("Evaluate called with empty formula");
                return formula;
            }

            if (!formula.Contains("{{"))
            {
                Debug.WriteLine($"Formula '{formula}' does not contain template expressions");
                return formula;
            }

            Debug.WriteLine($"Evaluating formula in cell {cell?.Address}: '{formula}'");

            try
            {
                // 식에서 모든 표현식 추출
                var expressions = ExpressionParser.ExtractExpressions(formula).ToList();
                Debug.WriteLine($"Extracted {expressions.Count} expressions: {string.Join(", ", expressions)}");

                if (expressions.Count == 0)
                {
                    Debug.WriteLine("No expressions found in formula");
                    return formula;
                }

                // 전체 식이 하나의 표현식인 경우
                if (expressions.Count == 1 && expressions[0] == formula)
                {
                    Debug.WriteLine("Formula is a single expression");
                    var parsedExpr = ExpressionParser.Parse(formula);
                    Debug.WriteLine($"Parsed expression: Type={parsedExpr.Type}, Var={parsedExpr.Variable}, Op={parsedExpr.Operation}");

                    switch (parsedExpr.Type)
                    {
                        case ExpressionType.Format:
                            Debug.WriteLine($"Processing format expression: {parsedExpr.Variable}:{parsedExpr.Operation}");
                            return EvaluateFormatExpression(parsedExpr, parameters);

                        case ExpressionType.Function:
                            Debug.WriteLine($"Processing function expression: {parsedExpr.Variable}|{parsedExpr.Operation}");
                            return EvaluateFunctionExpression(parsedExpr, cell, parameters);

                        default:
                            // 표준 표현식은 기본 평가자로 넘김
                            Debug.WriteLine("Processing standard expression");
                            try
                            {
                                var result = _baseEvaluator.Evaluate(formula, parameters);
                                Debug.WriteLine($"Base evaluator result: '{result}'");
                                return result;
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Base evaluator failed: {ex.Message}");
                                var result = ResolveStandardExpression(parsedExpr.Variable, parameters);
                                Debug.WriteLine($"Resolved standard expression: '{result}'");
                                return result;
                            }
                    }
                }

                // 여러 표현식이 있는 경우 각각 처리해서 결과 조합
                var finalResult = formula;
                Debug.WriteLine("Multiple expressions found, processing each");

                foreach (var expr in expressions)
                {
                    Debug.WriteLine($"Processing expression: '{expr}'");
                    var parsedExpr = ExpressionParser.Parse(expr);
                    Debug.WriteLine($"Parsed expression: Type={parsedExpr.Type}, Var={parsedExpr.Variable}, Op={parsedExpr.Operation}");

                    object value;

                    switch (parsedExpr.Type)
                    {
                        case ExpressionType.Format:
                            Debug.WriteLine($"Processing format expression: {parsedExpr.Variable}:{parsedExpr.Operation}");
                            value = EvaluateFormatExpression(parsedExpr, parameters);
                            Debug.WriteLine($"Format result: '{value}'");
                            break;

                        case ExpressionType.Function:
                            Debug.WriteLine($"Processing function expression: {parsedExpr.Variable}|{parsedExpr.Operation}");
                            value = EvaluateFunctionExpression(parsedExpr, cell, parameters);
                            Debug.WriteLine($"Function result: '{value}'");
                            break;

                        default:
                            Debug.WriteLine("Processing standard expression");
                            try
                            {
                                value = _baseEvaluator.Evaluate(expr, parameters);
                                Debug.WriteLine($"Base evaluator result: '{value}'");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Base evaluator failed: {ex.Message}");
                                value = ResolveStandardExpression(parsedExpr.Variable, parameters);
                                Debug.WriteLine($"Resolved standard expression: '{value}'");
                            }
                            break;
                    }

                    // 표현식을 값으로 대체
                    finalResult = finalResult.Replace(expr, value?.ToString() ?? string.Empty);
                    Debug.WriteLine($"Result after replacement: '{finalResult}'");
                }

                Debug.WriteLine($"Final evaluation result: '{finalResult}'");
                return finalResult;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error evaluating formula: {ex.Message}");
                Debug.WriteLine($"Exception stack trace: {ex.StackTrace}");

                // 오류 표시
                if (cell != null)
                {
                    try
                    {
                        var comment = cell.GetComment();
                        comment.AddText($"Error: {ex.Message}");
                    }
                    catch
                    {
                        // 주석 추가 실패 시 무시
                    }
                }

                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// 표준 변수 표현식 해석
        /// </summary>
        private object ResolveStandardExpression(string variableName, Parameter[] parameters)
        {
            // 매개변수에서 찾기
            foreach (var param in parameters.Where(p => p.ParameterExpression != null))
            {
                if (param.ParameterExpression.Name == variableName)
                    return param.Value;
            }

            // 템플릿에서 직접 해석
            var value = _template.ResolveVariable(variableName);
            if (value != null)
                return value;

            // 찾지 못한 경우 원래 표현식 반환
            return $"{{{{{variableName}}}}}";
        }

        /// <summary>
        /// 형식 표현식 ({{value:format}}) 평가
        /// </summary>
        private object EvaluateFormatExpression(ParsedExpression expression, Parameter[] parameters)
        {
            // 먼저 변수 부분 평가
            var varExpr = $"{{{{{expression.Variable}}}}}";
            Debug.WriteLine($"Evaluating variable part: '{varExpr}'");

            object value;

            try
            {
                value = _baseEvaluator.Evaluate(varExpr, parameters);
                Debug.WriteLine($"Base evaluator result for variable: '{value}'");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Base evaluator failed for variable: {ex.Message}");
                value = ResolveStandardExpression(expression.Variable, parameters);
                Debug.WriteLine($"Resolved variable value: '{value}'");
            }

            // 포맷터 찾기
            Debug.WriteLine($"Looking for formatter: '{expression.Operation}'");

            if (_template.TryGetFormatter(expression.Operation, out var formatter))
            {
                Debug.WriteLine($"Formatter found, applying to value: '{value}'");
                try
                {
                    var result = formatter(value, expression.Parameters);
                    Debug.WriteLine($"Formatter result: '{result}'");
                    return result;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Formatter application error: {ex.Message}");
                    Debug.WriteLine($"Exception stack trace: {ex.StackTrace}");
                    return value;
                }
            }

            Debug.WriteLine($"No custom formatter found, checking for standard formatting");

            // 포맷터가 없으면 .NET 표준 포맷팅 시도
            if (value is IFormattable formattable)
            {
                try
                {
                    var result = formattable.ToString(expression.Operation, null);
                    Debug.WriteLine($"Standard formatting result: '{result}'");
                    return result;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Standard formatting error: {ex.Message}");
                }
            }

            // 포맷터를 찾지 못한 경우 원래 값 반환
            Debug.WriteLine($"No formatter found, returning original value: '{value}'");
            return value;
        }

        /// <summary>
        /// 함수 표현식 ({{value|function}}) 평가
        /// </summary>
        private object EvaluateFunctionExpression(ParsedExpression expression, IXLCell cell, Parameter[] parameters)
        {
            // 먼저 변수 부분 평가
            var varExpr = $"{{{{{expression.Variable}}}}}";
            object value;

            try
            {
                value = _baseEvaluator.Evaluate(varExpr, parameters);
            }
            catch
            {
                value = ResolveStandardExpression(expression.Variable, parameters);
            }

            // 함수 찾기
            if (!_template.TryGetFunction(expression.Operation, out var function))
            {
                Debug.WriteLine($"Function not found: {expression.Operation}");
                return $"Unknown identifier '{expression.Operation}'";
            }

            // 함수 적용
            try
            {
                Debug.WriteLine($"Applying function {expression.Operation} to cell {cell.Address}");

                // 함수 처리에 사용할 임시 셀 생성
                var targetCell = cell is null
                    ? cell.Worksheet.Cell("A1")  // 기본 셀
                    : cell;                     // 제공된 셀

                // 함수 호출
                function(targetCell, value, expression.Parameters);

                // 셀 값 반환
                return targetCell.Value;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Function application error: {ex.Message}");
                return $"Error: {ex.Message}";
            }
        }
    }
}