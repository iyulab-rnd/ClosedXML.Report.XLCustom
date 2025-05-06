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
            Debug.WriteLine("XLCustomTemplate.Generate - Starting template generation");

            // 먼저 확장 표현식 처리
            // 기본 템플릿 엔진이 처리하기 전에 우리 표현식부터 처리
            ProcessEnhancedExpressionsBeforeGeneration();

            // 그 다음 기본 템플릿 엔진 실행
            var baseResult = _baseTemplate.Generate();
            Debug.WriteLine($"Base template generation completed with {baseResult.ParsingErrors.Count} errors");

            // 필요하다면 추가로 확장 표현식 처리
            ProcessEnhancedExpressions();

            // 오류 합치기
            foreach (var error in baseResult.ParsingErrors)
            {
                _errors.Add(error);
            }

            Debug.WriteLine($"Template generation completed with {_errors.Count} total errors");
            return new XLGenerateResult(_errors);
        }
        catch (Exception ex)
        {
            // 예상치 못한 예외 로깅
            Debug.WriteLine($"Unexpected error in Generate: {ex.Message}");
            Debug.WriteLine($"Exception stack trace: {ex.StackTrace}");
            _errors.Add(CreateTemplateError($"Unexpected error: {ex.Message}"));
            return new XLGenerateResult(_errors);
        }
    }

    /// <summary>
    /// Processes enhanced expressions before the base template generation
    /// </summary>
    private void ProcessEnhancedExpressionsBeforeGeneration()
    {
        Debug.WriteLine("Processing enhanced expressions before base generation");

        // 보이는 모든 워크시트 처리
        foreach (var ws in Workbook.Worksheets.Where(sh =>
            sh.Visibility == XLWorksheetVisibility.Visible &&
            !sh.PivotTables.Any()).ToArray())
        {
            Debug.WriteLine($"Pre-processing worksheet: {ws.Name}");

            // 확장 표현식(format, function)이 있는 셀 찾기
            var enhancedCells = ws.CellsUsed(c =>
                c.GetString().Contains("{{") &&
                (c.GetString().Contains(":") || c.GetString().Contains("|")));

            foreach (var cell in enhancedCells)
            {
                ProcessEnhancedExpressionBeforeGeneration(cell);
            }
        }
    }

    /// <summary>
    /// Processes a single enhanced expression before the base template generation
    /// </summary>
    private void ProcessEnhancedExpressionBeforeGeneration(IXLCell cell)
    {
        string value = cell.GetString();
        Debug.WriteLine($"Pre-processing cell {cell.Address}: Value = '{value}'");

        // 비표준 표현식만 찾기
        if (!ExpressionParser.IsEnhancedExpression(value))
        {
            Debug.WriteLine($"Cell {cell.Address} does not contain enhanced expressions");
            return;
        }

        try
        {
            // 표현식 파싱
            var expressions = ExpressionParser.ExtractExpressions(value).ToList();

            if (expressions.Count == 0)
                return;

            Debug.WriteLine($"Found {expressions.Count} expressions in cell {cell.Address}");

            // 표현식이 하나이고 셀 전체가 표현식인 경우
            if (expressions.Count == 1 && expressions[0] == value)
            {
                var parsedExpr = ExpressionParser.Parse(value);

                if (parsedExpr.Type == ExpressionType.Format)
                {
                    // 변수 값 직접 가져오기
                    var varValue = ResolveVariable(parsedExpr.Variable);

                    if (varValue != null)
                    {
                        Debug.WriteLine($"Resolved variable '{parsedExpr.Variable}' to '{varValue}'");

                        // 포맷터 적용
                        if (TryGetFormatter(parsedExpr.Operation, out var formatter))
                        {
                            var formattedValue = formatter(varValue, parsedExpr.Parameters);
                            Debug.WriteLine($"Applied formatter, result: '{formattedValue}'");

                            // 값 설정
                            cell.SetValue(formattedValue);
                            Debug.WriteLine($"Set cell {cell.Address} value to '{cell.Value}'");

                            // 처리 완료 표시
                            return;
                        }
                    }

                    // 변환에 실패한 경우 원래 표현식을 유지
                    Debug.WriteLine($"Failed to process format expression, keeping original");
                }
            }

            // 다른 경우는 표현식을 있는 그대로 유지
            Debug.WriteLine($"Expression will be processed normally: '{value}'");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error pre-processing cell {cell.Address}: {ex.Message}");
        }
    }

    /// <summary>
    /// Processes enhanced expressions after the base template generation
    /// </summary>
    private void ProcessEnhancedExpressions()
    {
        Debug.WriteLine("Processing enhanced expressions in all worksheets");

        // Process all visible worksheets that aren't pivot tables
        foreach (var ws in Workbook.Worksheets.Where(sh =>
            sh.Visibility == XLWorksheetVisibility.Visible &&
            !sh.PivotTables.Any()).ToArray())
        {
            Debug.WriteLine($"Processing worksheet: {ws.Name}");
            ProcessEnhancedExpressionsInWorksheet(ws);
        }

        Debug.WriteLine("Enhanced expressions processing completed");
    }

    /// <summary>
    /// Processes enhanced expressions in a specific worksheet
    /// </summary>
    private void ProcessEnhancedExpressionsInWorksheet(IXLWorksheet worksheet)
    {
        try
        {
            // Get a custom interpreter for this worksheet
            var interpreter = GetInterpreter(worksheet.Name);

            // Find cells with enhanced expressions (containing : or |)
            var enhancedCells = worksheet.CellsUsed(c =>
                c.GetString().Contains("{{") &&
                (c.GetString().Contains(":") || c.GetString().Contains("|")));

            if (!enhancedCells.Any())
                return;

            // Process each cell individually
            foreach (var cell in enhancedCells)
            {
                try
                {
                    ProcessEnhancedCell(cell, interpreter);
                }
                catch (Exception ex)
                {
                    // Log cell-specific errors but continue processing
                    _errors.Add(CreateTemplateError(
                        $"Error processing cell {cell.Address}: {ex.Message}",
                        worksheet.Range(cell.Address, cell.Address)));
                }
            }
        }
        catch (Exception ex)
        {
            // Log worksheet-level errors
            var dummyRange = worksheet.Range("A1", "A1");
            _errors.Add(CreateTemplateError(
                $"Error processing worksheet {worksheet.Name}: {ex.Message}",
                dummyRange));
        }
    }

    /// <summary>
    /// Processes a cell with enhanced expressions
    /// </summary>
    private void ProcessEnhancedCell(IXLCell cell, CustomRangeInterpreter interpreter)
    {
        string value = cell.GetString();

        // 디버그 정보 추가
        Debug.WriteLine($"ProcessEnhancedCell - Cell: {cell.Address}, Value: '{value}'");

        // Skip if not a template expression or already processed
        if (!value.Contains("{{") || (!value.Contains(":") && !value.Contains("|")))
        {
            Debug.WriteLine($"Cell {cell.Address} - Not an enhanced expression or already processed");
            return;
        }

        Debug.WriteLine($"Processing cell {cell.Address}: Current value = '{value}'");

        // Create custom evaluator
        var customEvaluator = new CustomFormulaEvaluator(this, interpreter.BaseEvaluator);

        try
        {
            // Process the expression
            Debug.WriteLine($"Evaluating expression: '{value}'");
            var result = customEvaluator.Evaluate(value, cell);
            Debug.WriteLine($"Expression evaluation result: '{result}'");

            // Set the result
            if (result != null)
            {
                if (value.StartsWith("&="))
                {
                    // Formula
                    Debug.WriteLine($"Setting formula: '{result}'");
                    cell.FormulaA1 = result.ToString();
                }
                else
                {
                    // Regular value
                    Debug.WriteLine($"Setting value: '{result}'");
                    cell.SetValue(value);
                }
            }

            Debug.WriteLine($"Cell {cell.Address} processed: Value = '{cell.Value}', Bold = {cell.Style.Font.Bold}");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error processing cell {cell.Address}: {ex.Message}");
            Debug.WriteLine($"Exception stack trace: {ex.StackTrace}");

            // Mark the cell with error
            cell.Value = $"Error: {ex.Message}";
            cell.Style.Font.FontColor = XLColor.Red;
        }
    }
}