namespace ClosedXML.Report.XLCustom
{
    public partial class XLCustomTemplate
    {
        // Track cells we've preprocessed
        private readonly HashSet<string> _preprocessedCells = new HashSet<string>();

        /// <summary>
        /// Processes collection metadata expressions like {{Collection.Count}}
        /// </summary>
        private void ProcessCollectionMetadata()
        {
            foreach (var worksheet in Workbook.Worksheets)
            {
                // 안전하게 사용된 셀의 배열 복사본을 먼저 생성
                var metadataCells = worksheet.CellsUsed(cell =>
                {
                    var value = cell.GetString();
                    return value.Contains("{{") && value.Contains(".Count") && value.Contains("}}");
                }).ToArray();

                foreach (var cell in metadataCells)
                {
                    string cellValue = cell.GetString();

                    // Simple regex to find collection metadata expressions
                    var matches = System.Text.RegularExpressions.Regex.Matches(
                        cellValue, @"\{\{([^{}\.]+)\.Count\}\}");

                    bool cellModified = false;

                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        string collectionName = match.Groups[1].Value;

                        // Try to get the collection from variables
                        if (TryGetCollection(collectionName, out var collection))
                        {
                            int count = collection.Count;

                            // Create variable name
                            string countVarName = $"{collectionName}_Count";

                            // Add variable if it doesn't exist
                            if (!_variables.ContainsKey(countVarName))
                            {
                                AddVariable(countVarName, count);
                            }

                            // Replace original expression with variable reference
                            cellValue = cellValue.Replace(match.Value, $"{{{{{countVarName}}}}}");
                            cellModified = true;
                        }
                    }

                    // Update cell if modified
                    if (cellModified)
                    {
                        cell.Value = cellValue;
                        MarkCellAsPreprocessed(cell);
                    }
                }
            }
        }

        /// <summary>
        /// Pre-processes enhanced expressions to make them compatible with the base engine
        /// </summary>
        private void PreprocessEnhancedExpressions()
        {
            // Process all visible worksheets
            foreach (var worksheet in Workbook.Worksheets.Where(sh =>
                sh.Visibility == XLWorksheetVisibility.Visible &&
                !sh.PivotTables.Any()).ToArray())
            {
                // Find cells with enhanced expressions - 안전하게 배열로 먼저 복사
                var enhancedCells = worksheet.CellsUsed(cell =>
                {
                    var value = cell.GetString();
                    return value.Contains("{{") && (value.Contains(":") || value.Contains("|"));
                }).ToArray();

                // Process each cell
                foreach (var cell in enhancedCells)
                {
                    PreprocessEnhancedCell(cell);
                }
            }

            // 모든 향상된 표현식을 처리한 후 인터프리터를 다시 초기화해야 함을 표시
            _interpreterInitialized = false;
        }

        /// <summary>
        /// Pre-process a cell with enhanced expressions by converting them to standard syntax
        /// </summary>
        private void PreprocessEnhancedCell(IXLCell cell)
        {
            string cellValue = cell.GetString();

            // Skip if not an enhanced expression
            if (!cellValue.Contains("{{") || (!cellValue.Contains(":") && !cellValue.Contains("|")))
                return;

            try
            {
                // Extract expressions
                var expressions = XLExpressionParser.ExtractExpressions(cellValue).ToArray(); // 안전하게 배열로 복사

                // Replace each enhanced expression
                bool cellModified = false;

                foreach (var expr in expressions)
                {
                    var parsedExpr = XLExpressionParser.Parse(expr);

                    // Skip standard expressions
                    if (parsedExpr.Type == XLExpressionType.Standard)
                        continue;

                    // Create a temporary variable to hold the enhanced expression result
                    string tempVarName = $"_temp_{Guid.NewGuid():N}";

                    // Add the expression logic to a preprocessing handler
                    AddEnhancedExpressionHandler(tempVarName, parsedExpr, cell);

                    // Replace the enhanced expression with a standard variable reference
                    cellValue = cellValue.Replace(expr, $"{{{{{tempVarName}}}}}");
                    cellModified = true;
                }

                // Update the cell if modified
                if (cellModified)
                {
                    cell.Value = cellValue;
                    MarkCellAsPreprocessed(cell);
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                _errors.Add(new TemplateError(
                    $"Error preprocessing cell {cell.Address.ToString()}: {ex.Message}",
                    cell.AsRange()));
            }
        }

        /// <summary>
        /// Post-processes enhanced expressions after the base template generation
        /// </summary>
        private void PostprocessEnhancedExpressions()
        {
            // Process any cells that might still have enhanced expressions
            foreach (var worksheet in Workbook.Worksheets.Where(sh =>
                sh.Visibility == XLWorksheetVisibility.Visible &&
                !sh.PivotTables.Any()).ToArray())
            {
                // Find cells with remaining enhanced expressions - 안전하게 배열로 복사
                var enhancedCells = worksheet.CellsUsed(cell =>
                {
                    var value = cell.GetString();
                    return value.Contains("{{") && (value.Contains(":") || value.Contains("|"));
                }).ToArray();

                // Process each cell
                foreach (var cell in enhancedCells)
                {
                    PostprocessEnhancedCell(cell);
                }
            }
        }

        /// <summary>
        /// Post-process a cell that still has enhanced expressions
        /// </summary>
        private void PostprocessEnhancedCell(IXLCell cell)
        {
            string cellValue = cell.GetString();

            // Skip if not an enhanced expression
            if (!cellValue.Contains("{{") || (!cellValue.Contains(":") && !cellValue.Contains("|")))
                return;

            try
            {
                // Create a custom evaluator
                var customEvaluator = new CustomFormulaEvaluator(this, _evaluator);

                // Evaluate the expression
                var result = customEvaluator.Evaluate(cellValue, cell);

                // Set the result
                if (result != null)
                {
                    if (cellValue.StartsWith("&="))
                    {
                        // Formula
                        cell.FormulaA1 = result.ToString();
                    }
                    else
                    {
                        // Regular value
                        cell.SetValue(XLCellValueConverter.FromObject(result));
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error
                _errors.Add(new TemplateError(
                    $"Error post-processing cell {cell.Address.ToString()}: {ex.Message}",
                    cell.AsRange()));

                // Mark cell with error
                cell.Value = $"Error: {ex.Message}";
                cell.Style.Font.FontColor = XLColor.Red;
            }
        }

        // Collection of enhanced expression handlers
        private readonly Dictionary<string, Action> _enhancedExpressionHandlers = new Dictionary<string, Action>();

        /// <summary>
        /// Add a handler for an enhanced expression
        /// </summary>
        private void AddEnhancedExpressionHandler(string tempVarName, ParsedExpression expression, IXLCell cell)
        {
            // Create a handler function
            Action handler = () =>
            {
                try
                {
                    // First, evaluate the variable part
                    string variableName = expression.Variable;
                    object value = null;

                    // Formula evaluation (e.g. item.Price * 1.1)
                    if (variableName.Contains(" * ") || variableName.Contains(" / ") ||
                        variableName.Contains(" + ") || variableName.Contains(" - "))
                    {
                        value = EvaluateFormula(variableName);
                    }
                    else
                    {
                        // Regular variable processing
                        value = ResolveVariable(variableName);
                    }

                    if (value == null)
                    {
                        // 값이 없더라도 최소한 빈 문자열을 할당하여 변수가 존재하도록 함
                        AddVariable(tempVarName, string.Empty);
                        return;
                    }

                    object result = value;

                    if (expression.Type == XLExpressionType.Format)
                    {
                        // Apply formatter
                        if (TryGetFormatter(expression.Operation, out var formatter))
                        {
                            result = formatter(value, expression.Parameters);
                        }
                        else if (value is IFormattable formattable)
                        {
                            // Try standard .NET formatting
                            try
                            {
                                result = formattable.ToString(expression.Operation, System.Globalization.CultureInfo.CurrentCulture);
                            }
                            catch
                            {
                                // Fall back to original value
                                result = value;
                            }
                        }
                    }
                    else if (expression.Type == XLExpressionType.Function && cell != null)
                    {
                        // Apply function
                        if (TryGetFunction(expression.Operation, out var function))
                        {
                            // Create a temporary cell to hold the result
                            var tempCell = cell.Worksheet.Cell(cell.Address.ToString());

                            // Apply function
                            function(tempCell, value, expression.Parameters);

                            // Get result
                            result = tempCell.Value;
                        }
                    }

                    // Add result as a template variable (will overwrite if exists)
                    AddVariable(tempVarName, result ?? string.Empty);
                }
                catch (Exception ex)
                {
                    // Log error
                    _errors.Add(new TemplateError($"Error processing enhanced expression: {ex.Message}", cell?.AsRange()));

                    // Add error as variable (not in interpreter yet)
                    AddVariable(tempVarName, $"Error: {ex.Message}");
                }
            };

            // Store handler
            _enhancedExpressionHandlers[tempVarName] = handler;

            // Execute handler immediately
            handler();
        }

        /// <summary>
        /// Evaluates a formula expression with variables
        /// </summary>
        public object EvaluateFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return null;

            try
            {
                // Simple arithmetic operations
                if (formula.Contains(" * "))
                {
                    return EvaluateOperation(formula, " * ", (a, b) => a * b);
                }
                else if (formula.Contains(" / "))
                {
                    return EvaluateOperation(formula, " / ", (a, b) => b != 0 ? a / b : 0);
                }
                else if (formula.Contains(" + "))
                {
                    return EvaluateOperation(formula, " + ", (a, b) => a + b);
                }
                else if (formula.Contains(" - "))
                {
                    return EvaluateOperation(formula, " - ", (a, b) => a - b);
                }

                // Single variable reference
                return ResolveVariable(formula);
            }
            catch (Exception ex)
            {
                _errors.Add(new TemplateError($"Formula evaluation error: {ex.Message}", null));
                return formula; // Return original on error
            }
        }

        private object EvaluateOperation(string formula, string operatorStr, Func<decimal, decimal, decimal> operation)
        {
            string[] parts = formula.Split(new[] { operatorStr }, StringSplitOptions.None);
            if (parts.Length != 2)
                return formula;

            string leftPart = parts[0].Trim();
            string rightPart = parts[1].Trim();

            // Evaluate left part
            object leftValue = null;
            if (leftPart.Contains("."))
            {
                leftValue = ResolveVariable(leftPart);
            }
            else if (decimal.TryParse(leftPart, out decimal leftConstant))
            {
                leftValue = leftConstant;
            }

            // Evaluate right part
            object rightValue = null;
            if (rightPart.Contains("."))
            {
                rightValue = ResolveVariable(rightPart);
            }
            else if (decimal.TryParse(rightPart, out decimal rightConstant))
            {
                rightValue = rightConstant;
            }

            // Perform calculation
            if (leftValue != null && rightValue != null)
            {
                if (decimal.TryParse(leftValue.ToString(), out decimal leftDecimal) &&
                    decimal.TryParse(rightValue.ToString(), out decimal rightDecimal))
                {
                    return operation(leftDecimal, rightDecimal);
                }

                // Handle string concatenation for + operator
                if (operatorStr == " + " && (leftValue is string || rightValue is string))
                {
                    return leftValue?.ToString() + rightValue?.ToString();
                }
            }

            return formula; // Return original if we can't handle it
        }

        private bool IsPreprocessedCell(IXLCell cell)
        {
            if (cell == null) return false;
            return _preprocessedCells.Contains($"{cell.Worksheet.Name}!{cell.Address.ToString()}");
        }

        private void MarkCellAsPreprocessed(IXLCell cell)
        {
            if (cell == null) return;
            _preprocessedCells.Add($"{cell.Worksheet.Name}!{cell.Address.ToString()}");
        }
    }
}