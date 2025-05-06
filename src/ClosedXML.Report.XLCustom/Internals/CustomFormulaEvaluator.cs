using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.XLCustom.Internals
{
    /// <summary>
    /// Enhanced formula evaluator for custom templates
    /// </summary>
    internal class CustomFormulaEvaluator
    {
        private readonly XLCustomTemplate _template;
        private readonly FormulaEvaluator _baseEvaluator;

        /// <summary>
        /// Initializes a new custom formula evaluator
        /// </summary>
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
                // Extract all expressions from the formula
                var expressions = XLExpressionParser.ExtractExpressions(formula).ToList();

                if (expressions.Count == 0)
                {
                    return formula;
                }

                // If the formula is a single expression
                if (expressions.Count == 1 && expressions[0] == formula)
                {
                    var parsedExpr = XLExpressionParser.Parse(formula);

                    switch (parsedExpr.Type)
                    {
                        case XLExpressionType.Format:
                            return EvaluateFormatExpression(parsedExpr, parameters);

                        case XLExpressionType.Function:
                            return EvaluateFunctionExpression(parsedExpr, cell, parameters);

                        default:
                            // Try standard expression first
                            try
                            {
                                if (_baseEvaluator != null)
                                {
                                    var result = _baseEvaluator.Evaluate(formula, parameters);
                                    return result;
                                }
                            }
                            catch (Exception)
                            {
                                // Fallback to variable resolution
                            }

                            // Try to resolve variable
                            var varResult = ResolveStandardExpression(parsedExpr.Variable, parameters);
                            return varResult;
                    }
                }

                // For multiple expressions, process each one
                var finalResult = formula;
                List<string> processedExpressions = new List<string>();

                foreach (var expr in expressions)
                {
                    // Skip expressions already processed to avoid infinite recursion
                    if (processedExpressions.Contains(expr))
                        continue;

                    processedExpressions.Add(expr);

                    var parsedExpr = XLExpressionParser.Parse(expr);
                    object value;

                    switch (parsedExpr.Type)
                    {
                        case XLExpressionType.Format:
                            value = EvaluateFormatExpression(parsedExpr, parameters);
                            break;

                        case XLExpressionType.Function:
                            value = EvaluateFunctionExpression(parsedExpr, cell, parameters);
                            break;

                        default:
                            try
                            {
                                if (_baseEvaluator != null)
                                {
                                    value = _baseEvaluator.Evaluate(expr, parameters);
                                }
                                else
                                {
                                    value = ResolveStandardExpression(parsedExpr.Variable, parameters);
                                }
                            }
                            catch (Exception)
                            {
                                value = ResolveStandardExpression(parsedExpr.Variable, parameters);
                            }
                            break;
                    }

                    // Replace expression with value
                    if (value != null)
                    {
                        finalResult = finalResult.Replace(expr, value.ToString() ?? string.Empty);
                    }
                }

                return finalResult;
            }
            catch (Exception ex)
            {
                // Display error
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
            // Check for arithmetic expression
            if (variableName.Contains(" * ") || variableName.Contains(" / ") ||
                variableName.Contains(" + ") || variableName.Contains(" - "))
            {
                var result = _template.EvaluateFormula(variableName);
                return result;
            }

            // Look in parameters first with exact match
            foreach (var param in parameters.Where(p => p.ParameterExpression != null))
            {
                if (string.Equals(param.ParameterExpression.Name, variableName, StringComparison.OrdinalIgnoreCase))
                {
                    return param.Value;
                }
            }

            // Resolve directly from template
            var value = _template.ResolveVariable(variableName);
            if (value != null)
            {
                return value;
            }

            // If not found, return original expression
            return $"{{{{{variableName}}}}}";
        }

        /// <summary>
        /// Evaluates format expression ({{value:format}})
        /// </summary>
        private object EvaluateFormatExpression(ParsedExpression expression, Parameter[] parameters)
        {
            // First evaluate the variable part
            var varExpr = $"{{{{{expression.Variable}}}}}";
            object value;

            try
            {
                if (_baseEvaluator != null)
                {
                    value = _baseEvaluator.Evaluate(varExpr, parameters);
                }
                else
                {
                    value = ResolveStandardExpression(expression.Variable, parameters);
                }
            }
            catch
            {
                value = ResolveStandardExpression(expression.Variable, parameters);
            }

            // Check if value is still an expression
            if (value is string strValue && strValue.StartsWith("{{") && strValue.EndsWith("}}"))
            {
                // Expression evaluation failed, try special handling

                // Try formula evaluation (e.g. item.Price * 1.1)
                if (expression.Variable.Contains(" * ") || expression.Variable.Contains(" / ") ||
                    expression.Variable.Contains(" + ") || expression.Variable.Contains(" - "))
                {
                    value = _template.EvaluateFormula(expression.Variable);
                }
            }

            // Look for formatter
            if (_template.TryGetFormatter(expression.Operation, out var formatter))
            {
                try
                {
                    var result = formatter(value, expression.Parameters);
                    return result;
                }
                catch (Exception)
                {
                    return value;
                }
            }

            // If no formatter found, try standard .NET formatting
            if (value is IFormattable formattable)
            {
                try
                {
                    var result = formattable.ToString(expression.Operation, CultureInfo.CurrentCulture);
                    return result;
                }
                catch (Exception)
                {
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
            if (cell == null)
                throw new ArgumentNullException(nameof(cell), "Cell is required for function execution");

            // First evaluate the variable part
            var varExpr = $"{{{{{expression.Variable}}}}}";
            object value;

            try
            {
                if (_baseEvaluator != null)
                {
                    value = _baseEvaluator.Evaluate(varExpr, parameters);
                }
                else
                {
                    value = ResolveStandardExpression(expression.Variable, parameters);
                }
            }
            catch
            {
                value = ResolveStandardExpression(expression.Variable, parameters);
            }

            // Check if value is still an expression
            if (value is string strValue && strValue.StartsWith("{{") && strValue.EndsWith("}}"))
            {
                // Expression evaluation failed, try special handling

                // Try formula evaluation (e.g. item.Price * 1.1)
                if (expression.Variable.Contains(" * ") || expression.Variable.Contains(" / ") ||
                    expression.Variable.Contains(" + ") || expression.Variable.Contains(" - "))
                {
                    value = _template.EvaluateFormula(expression.Variable);
                }
            }

            // Look for function
            if (!_template.TryGetFunction(expression.Operation, out var function))
            {
                return $"Unknown function '{expression.Operation}'";
            }

            // Apply function
            try
            {
                // Clone the cell to avoid modifying the original
                var tempCell = cell;

                // Call function
                function(tempCell, value, expression.Parameters);

                // Return cell value
                return tempCell.Value;
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
}