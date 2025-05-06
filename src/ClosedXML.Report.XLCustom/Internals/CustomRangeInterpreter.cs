namespace ClosedXML.Report.XLCustom.Internals;

/// <summary>
/// Custom range interpreter that extends the standard RangeInterpreter
/// </summary>
internal class CustomRangeInterpreter : RangeInterpreter
{
    private readonly XLCustomTemplate _template;
    private readonly TemplateErrors _errors;
    private FormulaEvaluator _baseEvaluator;

    /// <summary>
    /// Gets the base formula evaluator
    /// </summary>
    public FormulaEvaluator BaseEvaluator
    {
        get
        {
            if (_baseEvaluator == null)
            {
                // 리플렉션으로 기본 평가자 가져오기
                var fieldInfo = typeof(RangeInterpreter).GetField("_evaluator",
                    BindingFlags.NonPublic | BindingFlags.Instance);

                if (fieldInfo != null)
                {
                    _baseEvaluator = (FormulaEvaluator)fieldInfo.GetValue(this);
                }
                else
                {
                    throw new InvalidOperationException(
                        "Could not access base evaluator. Library may be incompatible with this version of ClosedXML.Report.");
                }
            }

            return _baseEvaluator;
        }
    }

    public CustomRangeInterpreter(string alias, TemplateErrors errors, XLCustomTemplate template)
        : base(alias, errors)
    {
        _template = template ?? throw new ArgumentNullException(nameof(template));
        _errors = errors ?? throw new ArgumentNullException(nameof(errors));
    }

    // 값 평가 오버라이드 - 핵심 기능
    public override void EvaluateValues(IXLRange range, params Parameter[] pars)
    {
        Debug.WriteLine($"CustomRangeInterpreter.EvaluateValues: 범위 {range.RangeAddress}에서 확장 표현식 평가");

        try
        {
            // 변수가 글로벌 리졸버에 있는지 먼저 확인해서 추가
            foreach (var par in pars)
            {
                if (_template.TryResolveGlobal(par.ParameterExpression.Name, out var value))
                {
                    AddVariable(par.ParameterExpression.Name, value);
                }
            }

            // 기본 구현 호출 - 일반적인 ClosedXML.Report 표현식 처리
            base.EvaluateValues(range, pars);

            // 확장 표현식 처리
            var customEvaluator = new CustomFormulaEvaluator(_template, BaseEvaluator);

            // 확장 표현식이 있는 셀 검색
            var enhancedCells = range.CellsUsed(cell =>
            {
                var value = cell.GetString();
                return value.Contains("{{") && (value.Contains(":") || value.Contains("|"));
            });

            foreach (var cell in enhancedCells)
            {
                string cellValue = cell.GetString();

                // 셀이 여전히 확장 표현식을 포함하는지 확인
                // (기본 처리 후에도 남아있는지)
                if (cellValue.Contains("{{") && (cellValue.Contains(":") || cellValue.Contains("|")))
                {
                    Debug.WriteLine($"Processing enhanced expression in cell {cell.Address}: {cellValue}");

                    try
                    {
                        // 표현식 평가
                        var result = customEvaluator.Evaluate(cellValue, cell, pars);

                        // 값 설정
                        if (result != null)
                        {
                            if (cellValue.StartsWith("&="))
                            {
                                // 수식은 FormulaA1로 설정
                                cell.FormulaA1 = result.ToString();
                            }
                            else
                            {
                                // 일반 값은 Value로 설정
                                cell.SetValue(result);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error processing cell {cell.Address}: {ex.Message}");

                        // 오류 표시
                        _errors.Add(new TemplateError(
                            $"Error processing enhanced expression in cell {cell.Address}: {ex.Message}",
                            range.Worksheet.Range(cell.Address, cell.Address)));

                        cell.Value = $"Error: {ex.Message}";
                        cell.Style.Font.FontColor = XLColor.Red;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error in EvaluateValues: {ex.Message}");
            _errors.Add(new TemplateError($"Error evaluating range: {ex.Message}", range));
        }
    }
}