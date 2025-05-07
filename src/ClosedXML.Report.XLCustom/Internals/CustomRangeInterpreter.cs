namespace ClosedXML.Report.XLCustom.Internals
{
    /// <summary>
    /// Custom range interpreter that extends the base RangeInterpreter to use the custom formula evaluator
    /// </summary>
    internal class CustomRangeInterpreter : RangeInterpreter
    {
        private readonly Dictionary<string, FormatHandler> _formatters;
        private readonly Dictionary<string, FunctionHandler> _functions;
        private readonly CustomFormulaEvaluator _customEvaluator;

        // 오류 정보 접근을 위한 속성
        public TemplateErrors Errors { get; }

        public CustomRangeInterpreter(string alias, TemplateErrors errors,
                                     Dictionary<string, FormatHandler> formatters,
                                     Dictionary<string, FunctionHandler> functions)
            : base(alias, errors)
        {
            _formatters = formatters;
            _functions = functions;
            Errors = errors;

            // 부모의 FormulaEvaluator를 확장한 커스텀 평가기 생성
            _customEvaluator = new CustomFormulaEvaluator(
                GetBaseEvaluator(), _formatters, _functions);
        }

        // RangeInterpreter의 메서드 오버라이드
        public override void EvaluateValues(IXLRange range, params Parameter[] pars)
        {
            // 부모 구현 활용
            base.EvaluateValues(range, pars);

            // 커스텀 표현식 처리
            ProcessCustomExpressions(range, pars);
        }

        // 부모의 FormulaEvaluator 접근
        private FormulaEvaluator GetBaseEvaluator()
        {
            // 필요한 경우 리플렉션 사용
            var field = typeof(RangeInterpreter).GetField("_evaluator",
                BindingFlags.NonPublic | BindingFlags.Instance);

            if (field == null)
                throw new InvalidOperationException("Cannot access base evaluator");

            return (FormulaEvaluator)field.GetValue(this);
        }

        // 커스텀 표현식 처리
        private void ProcessCustomExpressions(IXLRange range, Parameter[] pars)
        {
            // 범위 내에서 커스텀 표현식이 있는 셀 찾기
            var cellsToProcess = FindCellsWithCustomExpressions(range);

            foreach (var cell in cellsToProcess)
            {
                ProcessCustomCell(cell, pars);
            }
        }

        // 커스텀 표현식이 있는 셀 찾기
        private IEnumerable<IXLCell> FindCellsWithCustomExpressions(IXLRange range)
        {
            var innerRanges = GetInnerRanges(range);

            return range.CellsUsed()
                .Where(c => !c.HasFormula && HasCustomSyntax(c.GetString()) &&
                      !innerRanges.Any(nr => nr.Ranges.Contains(c.AsRange())));
        }

        // 커스텀 표현식 확인
        private bool HasCustomSyntax(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;

            // 포맷 표현식: {{value:format}}
            if (value.Contains("{{") && value.Contains("}}") && value.Contains(":"))
                return true;

            // 함수 표현식: {{value|function}}
            if (value.Contains("{{") && value.Contains("}}") && value.Contains("|"))
                return true;

            return false;
        }

        // 셀 처리
        private void ProcessCustomCell(IXLCell cell, Parameter[] pars)
        {
            var value = cell.GetString();

            try
            {
                // 커스텀 평가기로 표현식 처리
                var result = _customEvaluator.Evaluate(value, cell, pars);

                // 함수 표현식이 아닌 경우에만 값 설정
                // (함수 표현식은 셀에 직접 작업할 수 있음)
                if (!value.Contains("|"))
                {
                    cell.SetValue(XLCellValueConverter.FromObject(result));
                }
            }
            catch (Exception ex)
            {
                // 오류 처리
                cell.Value = $"Error: {ex.Message}";
                cell.Style.Font.FontColor = XLColor.Red;
                Errors.Add(new TemplateError(ex.Message, cell.AsRange()));
            }
        }

        // 범위 내의 정의된 이름 목록 가져오기
        private IEnumerable<IXLDefinedName> GetInnerRanges(IXLRange range)
        {
            return range.GetContainingNames();
        }

        // 글로벌 리졸버 설정
        public void SetGlobalResolver(GlobalVariableResolver resolver)
        {
            _customEvaluator.SetGlobalResolver(resolver);
        }
    }
}