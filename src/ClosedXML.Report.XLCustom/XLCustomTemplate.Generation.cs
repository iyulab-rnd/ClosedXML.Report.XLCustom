namespace ClosedXML.Report.XLCustom;

public partial class XLCustomTemplate
{
    /// <summary>
    /// Generates the report by processing all template expressions
    /// </summary>
    public XLGenerateResult Generate()
    {
        CheckIsDisposed();

        try
        {
            // 컬렉션 메타데이터 처리
            ProcessCollectionMetadata();

            // 디버깅을 위한 변수 로깅
            LogVariables();

            // Preprocess enhanced expressions to make them compatible with base engine
            PreprocessEnhancedExpressions();

            // 다시 로깅
            LogVariables();

            // Run the base template engine
            var baseResult = _baseTemplate.Generate();

            // Post-process any remaining enhanced expressions
            PostprocessEnhancedExpressions();

            // Merge errors
            foreach (var error in baseResult.ParsingErrors)
            {
                // Skip syntax errors on cells we've already processed
                if (error.Message.Contains("Syntax error") &&
                    error.Range != null &&
                    IsPreprocessedCell(error.Range.FirstCell()))
                {
                    continue;
                }

                _errors.Add(error);
            }

            return new XLGenerateResult(_errors);
        }
        catch (Exception ex)
        {
            _errors.Add(CreateTemplateError($"Unexpected error: {ex.Message}"));
            return new XLGenerateResult(_errors);
        }
    }

    // Track cells we've preprocessed
    private readonly HashSet<string> _preprocessedCells = new HashSet<string>();

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

    /// <summary>
    /// Logs variables for debugging
    /// </summary>
    private void LogVariables()
    {
        Debug.WriteLine("=== Variables in Template ===");
        foreach (var variable in _variables)
        {
            Debug.WriteLine($"  {variable.Key} = {variable.Value}");
        }

        Debug.WriteLine("=== Variables in Base Template ===");
        foreach (var variable in _variablesAddedToBase)
        {
            Debug.WriteLine($"  {variable}");
        }
        Debug.WriteLine("============================");
    }

    /// <summary>
    /// Processes collection metadata expressions like {{Collection.Count}}
    /// </summary>
    private void ProcessCollectionMetadata()
    {
        Debug.WriteLine("Processing collection metadata...");

        foreach (var worksheet in Workbook.Worksheets)
        {
            var metadataCells = worksheet.CellsUsed(cell =>
            {
                var value = cell.GetString();
                return value.Contains("{{") && value.Contains(".Count") && value.Contains("}}");
            });

            foreach (var cell in metadataCells)
            {
                string cellValue = cell.GetString();
                Debug.WriteLine($"Processing metadata cell {cell.Address}: {cellValue}");

                var matches = System.Text.RegularExpressions.Regex.Matches(
                    cellValue, @"\{\{([^{}\.]+)\.Count\}\}");

                bool cellModified = false;

                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    string collectionName = match.Groups[1].Value;
                    Debug.WriteLine($"Found collection metadata: {collectionName}.Count");

                    if (TryGetCollection(collectionName, out var collection))
                    {
                        int count = collection.Count;
                        Debug.WriteLine($"Collection {collectionName} has {count} items");

                        // 변수명 생성
                        string countVarName = $"{collectionName}_Count";

                        // 변수가 이미 존재하는지 확인 후 추가
                        if (!_variables.ContainsKey(countVarName))
                        {
                            AddVariable(countVarName, count);
                        }
                        else
                        {
                            Debug.WriteLine($"Variable {countVarName} already exists, skipping");
                        }

                        // 원본 표현식을 변수 참조로 변경
                        cellValue = cellValue.Replace(match.Value, $"{{{{{countVarName}}}}}");
                        cellModified = true;
                    }
                    else
                    {
                        Debug.WriteLine($"Collection {collectionName} not found");
                    }
                }

                // Update cell if modified
                if (cellModified)
                {
                    Debug.WriteLine($"Updating cell {cell.Address} with value: {cellValue}");
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
            // Find cells with enhanced expressions
            var enhancedCells = worksheet.CellsUsed(cell =>
            {
                var value = cell.GetString();
                return value.Contains("{{") && (value.Contains(":") || value.Contains("|"));
            });

            // Process each cell
            foreach (var cell in enhancedCells)
            {
                PreprocessEnhancedCell(cell);
            }
        }
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
            Debug.WriteLine($"Preprocessing cell {cell.Address} with value: {cellValue}");

            // Extract expressions
            var expressions = ExpressionParser.ExtractExpressions(cellValue).ToList();

            // Replace each enhanced expression
            bool cellModified = false;

            foreach (var expr in expressions)
            {
                var parsedExpr = ExpressionParser.Parse(expr);

                // Skip standard expressions
                if (parsedExpr.Type == ExpressionType.Standard)
                    continue;

                // Create a temporary variable to hold the enhanced expression result
                string tempVarName = $"_temp_{Guid.NewGuid():N}";
                Debug.WriteLine($"Created temp variable: {tempVarName} for expression: {expr}");

                // Add the expression logic to a preprocessing handler
                AddEnhancedExpressionHandler(tempVarName, parsedExpr, cell);

                // Replace the enhanced expression with a standard variable reference
                cellValue = cellValue.Replace(expr, $"{{{{{tempVarName}}}}}");
                cellModified = true;
            }

            // Update the cell if modified
            if (cellModified)
            {
                Debug.WriteLine($"Cell modified. New value: {cellValue}");
                cell.Value = cellValue;
                MarkCellAsPreprocessed(cell);
            }
        }
        catch (Exception ex)
        {
            // Log error but continue
            Debug.WriteLine($"Error preprocessing cell {cell.Address}: {ex.Message}");
            _errors.Add(new TemplateError(
                $"Error preprocessing cell {cell.Address.ToString()}: {ex.Message}",
                cell.AsRange()));
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
        Action handler = () => {
            try
            {
                Debug.WriteLine($"Processing enhanced expression handler for {tempVarName}, expression: {expression.OriginalExpression}");

                // First, evaluate the variable part
                string variableName = expression.Variable;
                Debug.WriteLine($"Evaluating variable: {variableName}");

                object value = null;

                // 수식 평가 (예: item.Price * 1.1)
                if (variableName.Contains(" * ") || variableName.Contains(" / ") ||
                    variableName.Contains(" + ") || variableName.Contains(" - "))
                {
                    value = EvaluateFormula(variableName);
                    Debug.WriteLine($"Formula evaluation result: {value}");
                }
                else
                {
                    // 일반 변수 처리
                    value = ResolveVariable(variableName);
                    Debug.WriteLine($"Variable resolution result: {value}");
                }

                if (value == null)
                {
                    Debug.WriteLine($"Warning: Could not resolve variable {variableName}");
                    return; // Can't process without a value
                }

                object result = value;

                if (expression.Type == ExpressionType.Format)
                {
                    // Apply formatter
                    Debug.WriteLine($"Applying formatter: {expression.Operation} to value: {value}");
                    if (TryGetFormatter(expression.Operation, out var formatter))
                    {
                        result = formatter(value, expression.Parameters);
                        Debug.WriteLine($"Formatter result: {result}");
                    }
                    else if (value is IFormattable formattable)
                    {
                        // Try standard .NET formatting
                        try
                        {
                            result = formattable.ToString(expression.Operation, System.Globalization.CultureInfo.CurrentCulture);
                            Debug.WriteLine($"Standard formatting result: {result}");
                        }
                        catch (Exception ex)
                        {
                            // Fall back to original value
                            Debug.WriteLine($"Formatting failed: {ex.Message}");
                            result = value;
                        }
                    }
                }
                else if (expression.Type == ExpressionType.Function && cell != null)
                {
                    // Apply function
                    Debug.WriteLine($"Applying function: {expression.Operation} to value: {value}");
                    if (TryGetFunction(expression.Operation, out var function))
                    {
                        // Create a temporary cell to hold the result
                        var tempCell = cell.Worksheet.Cell(cell.Address.ToString());

                        // Apply function
                        function(tempCell, value, expression.Parameters);

                        // Get result
                        result = tempCell.Value;
                        Debug.WriteLine($"Function result: {result}");
                    }
                }

                // Add result as a template variable
                Debug.WriteLine($"Adding temp variable {tempVarName} with value: {result}");
                AddVariable(tempVarName, result);

                // Always add to base template
                _baseTemplate.AddVariable(tempVarName, result);
                _variablesAddedToBase.Add(tempVarName);
            }
            catch (Exception ex)
            {
                // Log error
                Debug.WriteLine($"Error in enhanced expression handler: {ex.Message}");
                _errors.Add($"Error processing enhanced expression: {ex.Message}");

                // Add error as variable
                AddVariable(tempVarName, $"Error: {ex.Message}");

                // Add to base template too
                _baseTemplate.AddVariable(tempVarName, $"Error: {ex.Message}");
                _variablesAddedToBase.Add(tempVarName);
            }
        };

        // Store handler
        _enhancedExpressionHandlers[tempVarName] = handler;

        // Execute handler immediately
        handler();
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
            // Find cells with remaining enhanced expressions
            var enhancedCells = worksheet.CellsUsed(cell =>
            {
                var value = cell.GetString();
                return value.Contains("{{") && (value.Contains(":") || value.Contains("|"));
            });

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
            Debug.WriteLine($"Post-processing cell {cell.Address} with value: {cellValue}");

            // Create a custom evaluator
            var customEvaluator = new CustomFormulaEvaluator(this, null);

            // Evaluate the expression
            var result = customEvaluator.Evaluate(cellValue, cell);
            Debug.WriteLine($"Evaluation result: {result}");

            // Set the result
            if (result != null)
            {
                if (cellValue.StartsWith("&="))
                {
                    // Formula
                    cell.FormulaA1 = result.ToString();
                    Debug.WriteLine($"Set formula: {result}");
                }
                else
                {
                    // Regular value
                    cell.SetValue(result);
                    Debug.WriteLine($"Set value: {result}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log error
            Debug.WriteLine($"Error post-processing cell {cell.Address}: {ex.Message}");
            _errors.Add(new TemplateError(
                $"Error post-processing cell {cell.Address.ToString()}: {ex.Message}",
                cell.AsRange()));

            // Mark cell with error
            cell.Value = $"Error: {ex.Message}";
            cell.Style.Font.FontColor = XLColor.Red;
        }
    }
}