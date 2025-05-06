namespace ClosedXML.Report.XLCustom.Internals;

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
        _baseEvaluator = baseEvaluator;
    }

    /// <summary>
    /// Evaluates a formula, including enhanced expressions
    /// </summary>
    public object Evaluate(string formula, IXLCell cell, params Parameter[] parameters)
    {
        if (string.IsNullOrEmpty(formula))
        {
            return formula;
        }

        if (!formula.Contains("{{"))
        {
            return formula;
        }

        try
        {
            Debug.WriteLine($"Evaluating formula: {formula} for cell {cell?.Address}");

            // Extract all expressions from the formula
            var expressions = ExpressionParser.ExtractExpressions(formula).ToList();

            if (expressions.Count == 0)
            {
                return formula;
            }

            // If the formula is a single expression
            if (expressions.Count == 1 && expressions[0] == formula)
            {
                var parsedExpr = ExpressionParser.Parse(formula);
                Debug.WriteLine($"Parsed expression: Type={parsedExpr.Type}, Variable={parsedExpr.Variable}, Operation={parsedExpr.Operation}");

                switch (parsedExpr.Type)
                {
                    case ExpressionType.Format:
                        return EvaluateFormatExpression(parsedExpr, parameters);

                    case ExpressionType.Function:
                        return EvaluateFunctionExpression(parsedExpr, cell, parameters);

                    default:
                        // 표준 표현식 먼저 시도
                        try
                        {
                            if (_baseEvaluator != null)
                            {
                                var result = _baseEvaluator.Evaluate(formula, parameters);
                                Debug.WriteLine($"Base evaluator result: {result}");
                                return result;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Base evaluator failed: {ex.Message}");
                        }

                        // 변수 해결 시도
                        var varResult = ResolveStandardExpression(parsedExpr.Variable, parameters);
                        Debug.WriteLine($"Variable resolution result: {varResult}");
                        return varResult;
                }
            }

            // For multiple expressions, process each one
            var finalResult = formula;
            Debug.WriteLine($"Processing multiple expressions in formula");

            foreach (var expr in expressions)
            {
                var parsedExpr = ExpressionParser.Parse(expr);
                object value;

                switch (parsedExpr.Type)
                {
                    case ExpressionType.Format:
                        value = EvaluateFormatExpression(parsedExpr, parameters);
                        Debug.WriteLine($"Format expression result: {value}");
                        break;

                    case ExpressionType.Function:
                        value = EvaluateFunctionExpression(parsedExpr, cell, parameters);
                        Debug.WriteLine($"Function expression result: {value}");
                        break;

                    default:
                        try
                        {
                            if (_baseEvaluator != null)
                            {
                                value = _baseEvaluator.Evaluate(expr, parameters);
                                Debug.WriteLine($"Base evaluator result: {value}");
                            }
                            else
                            {
                                value = ResolveStandardExpression(parsedExpr.Variable, parameters);
                                Debug.WriteLine($"Variable resolution result: {value}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Evaluation failed: {ex.Message}");
                            value = ResolveStandardExpression(parsedExpr.Variable, parameters);
                        }
                        break;
                }

                // Replace expression with value
                if (value != null)
                {
                    finalResult = finalResult.Replace(expr, value.ToString() ?? string.Empty);
                    Debug.WriteLine($"Replaced expression. New result: {finalResult}");
                }
            }

            return finalResult;
        }
        catch (Exception ex)
        {
            // Display error
            Debug.WriteLine($"Error evaluating formula: {ex.Message}");
            if (cell != null)
            {
                try
                {
                    var comment = cell.GetComment();
                    comment.AddText($"Error: {ex.Message}");
                }
                catch
                {
                    // Ignore comment failure
                }
            }

            return $"Error: {ex.Message}";
        }
    }

    /// <summary>
    /// Resolves standard variable expression
    /// </summary>
    private object ResolveStandardExpression(string variableName, Parameter[] parameters)
    {
        Debug.WriteLine($"Resolving standard expression: {variableName}");

        // 수식에 '*' 이 포함된 경우 (예: item.Price * 1.1)
        if (variableName.Contains(" * ") || variableName.Contains(" / ") ||
            variableName.Contains(" + ") || variableName.Contains(" - "))
        {
            var result = _template.EvaluateFormula(variableName);
            Debug.WriteLine($"Formula evaluation result: {result}");
            return result;
        }

        // Look in parameters
        foreach (var param in parameters.Where(p => p.ParameterExpression != null))
        {
            if (param.ParameterExpression.Name == variableName)
            {
                Debug.WriteLine($"Found in parameters: {param.Value}");
                return param.Value;
            }
        }

        // Resolve directly from template
        var value = _template.ResolveVariable(variableName);
        if (value != null)
        {
            Debug.WriteLine($"Resolved from template: {value}");
            return value;
        }

        // If not found, return original expression
        Debug.WriteLine($"Could not resolve: {variableName}");
        return $"{{{{{variableName}}}}}";
    }

    /// <summary>
    /// Evaluates format expression ({{value:format}})
    /// </summary>
    private object EvaluateFormatExpression(ParsedExpression expression, Parameter[] parameters)
    {
        Debug.WriteLine($"Evaluating format expression: {expression.OriginalExpression}");

        // First evaluate the variable part
        var varExpr = $"{{{{{expression.Variable}}}}}";
        Debug.WriteLine($"Variable expression: {varExpr}");

        object value;

        try
        {
            if (_baseEvaluator != null)
            {
                value = _baseEvaluator.Evaluate(varExpr, parameters);
                Debug.WriteLine($"Base evaluator result: {value}");
            }
            else
            {
                value = ResolveStandardExpression(expression.Variable, parameters);
                Debug.WriteLine($"Standard expression resolution: {value}");
            }
        }
        catch
        {
            value = ResolveStandardExpression(expression.Variable, parameters);
            Debug.WriteLine($"Fallback resolution: {value}");
        }

        // 값이 여전히 표현식 형태인지 확인
        if (value is string strValue && strValue.StartsWith("{{") && strValue.EndsWith("}}"))
        {
            // 표현식 평가 실패, 특별 처리 시도

            // 수식 평가 (item.Price * 1.1 같은 형태)
            if (expression.Variable.Contains(" * ") || expression.Variable.Contains(" / ") ||
                expression.Variable.Contains(" + ") || expression.Variable.Contains(" - "))
            {
                value = _template.EvaluateFormula(expression.Variable);
                Debug.WriteLine($"Formula evaluation result: {value}");
            }
        }

        // Look for formatter
        if (_template.TryGetFormatter(expression.Operation, out var formatter))
        {
            try
            {
                Debug.WriteLine($"Applying formatter: {expression.Operation}");
                var result = formatter(value, expression.Parameters);
                Debug.WriteLine($"Formatter result: {result}");
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Formatter error: {ex.Message}");
                return value;
            }
        }

        // If no formatter found, try standard .NET formatting
        if (value is IFormattable formattable)
        {
            try
            {
                Debug.WriteLine($"Applying standard format: {expression.Operation}");
                var result = formattable.ToString(expression.Operation, System.Globalization.CultureInfo.CurrentCulture);
                Debug.WriteLine($"Formatting result: {result}");
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Format error: {ex.Message}");
                // Fall through if format fails
            }
        }

        // If no formatter found, return original value
        return value;
    }

    /// <summary>
    /// Evaluates function expression ({{value|function}})
    /// </summary>
    private object EvaluateFunctionExpression(ParsedExpression expression, IXLCell cell, Parameter[] parameters)
    {
        Debug.WriteLine($"Evaluating function expression: {expression.OriginalExpression}");

        // First evaluate the variable part
        var varExpr = $"{{{{{expression.Variable}}}}}";
        object value;

        try
        {
            if (_baseEvaluator != null)
            {
                value = _baseEvaluator.Evaluate(varExpr, parameters);
                Debug.WriteLine($"Base evaluator result: {value}");
            }
            else
            {
                value = ResolveStandardExpression(expression.Variable, parameters);
                Debug.WriteLine($"Standard expression resolution: {value}");
            }
        }
        catch
        {
            value = ResolveStandardExpression(expression.Variable, parameters);
            Debug.WriteLine($"Fallback resolution: {value}");
        }

        // 값이 여전히 표현식 형태인지 확인
        if (value is string strValue && strValue.StartsWith("{{") && strValue.EndsWith("}}"))
        {
            // 표현식 평가 실패, 특별 처리 시도

            // 수식 평가 (item.Price * 1.1 같은 형태)
            if (expression.Variable.Contains(" * ") || expression.Variable.Contains(" / ") ||
                expression.Variable.Contains(" + ") || expression.Variable.Contains(" - "))
            {
                value = _template.EvaluateFormula(expression.Variable);
                Debug.WriteLine($"Formula evaluation result: {value}");
            }
        }

        // Look for function
        if (!_template.TryGetFunction(expression.Operation, out var function))
        {
            Debug.WriteLine($"Function not found: {expression.Operation}");
            return $"Unknown function '{expression.Operation}'";
        }

        // Apply function
        try
        {
            Debug.WriteLine($"Applying function: {expression.Operation}");

            // Create target cell
            var targetCell = cell is null
                ? null  // We can't create a default cell safely
                : cell;  // Provided cell

            if (targetCell == null)
            {
                Debug.WriteLine("No target cell provided");
                return value;
            }

            // Call function
            function(targetCell, value, expression.Parameters);
            Debug.WriteLine($"Function applied. Cell value: {targetCell.Value}");

            // Return cell value
            return targetCell.Value;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Function error: {ex.Message}");
            return $"Error: {ex.Message}";
        }
    }
}