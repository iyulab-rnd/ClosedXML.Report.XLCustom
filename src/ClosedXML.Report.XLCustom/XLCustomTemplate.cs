using ClosedXML.Report;

namespace ClosedXML.Report.XLCustom
{
    public partial class XLCustomTemplate : IXLTemplate
    {
        private readonly IXLWorkbook _workbook;
        private readonly CustomRangeInterpreter _interpreter;
        private readonly Dictionary<string, FormatHandler> _formatters = new();
        private readonly Dictionary<string, FunctionHandler> _functions = new();
        private bool _isDisposed;

        public IXLWorkbook Workbook => _workbook;
        public bool IsDisposed => _isDisposed;

        public XLCustomTemplate(string fileName)
        {
            _workbook = new XLWorkbook(fileName);
            _interpreter = new CustomRangeInterpreter("", new TemplateErrors(), _formatters, _functions);
        }

        public XLCustomTemplate(Stream stream)
        {
            _workbook = new XLWorkbook(stream);
            _interpreter = new CustomRangeInterpreter("", new TemplateErrors(), _formatters, _functions);
        }

        public XLCustomTemplate(IXLWorkbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _interpreter = new CustomRangeInterpreter("", new TemplateErrors(), _formatters, _functions);
        }

        public XLGenerateResult Generate()
        {
            if (_isDisposed)
                throw new ObjectDisposedException("Template has been disposed");

            // 워크시트 컬렉션의 복사본을 생성하여 순회
            foreach (var ws in _workbook.Worksheets.Where(sh =>
                sh.Visibility == XLWorksheetVisibility.Visible &&
                !sh.PivotTables.Any()).ToArray())
            {
                ws.ReplaceCFFormulaeToR1C1();
                _interpreter.Evaluate(ws.AsRange());
                ws.ReplaceCFFormulaeToA1();
            }

            return new XLGenerateResult(_interpreter.Errors);
        }

        public void AddVariable(object value) => _interpreter.AddObjectVariable(value);

        public void AddVariable(string alias, object value) => _interpreter.AddVariable(alias, value);

        // 저장 관련 메서드들
        public void SaveAs(string file) => _workbook.SaveAs(file);
        public void SaveAs(string file, SaveOptions options) => _workbook.SaveAs(file, options);
        public void SaveAs(string file, bool validate, bool evaluateFormulae = false) => _workbook.SaveAs(file, validate, evaluateFormulae);
        public void SaveAs(Stream stream) => _workbook.SaveAs(stream);
        public void SaveAs(Stream stream, SaveOptions options) => _workbook.SaveAs(stream, options);
        public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false) => _workbook.SaveAs(stream, validate, evaluateFormulae);

        // 확장 메서드: 포맷터 등록
        public XLCustomTemplate RegisterFormat(string name, FormatHandler handler)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name));

            _formatters[name.ToLowerInvariant()] = handler ??
                throw new ArgumentNullException(nameof(handler));

            return this;
        }

        // 확장 메서드: 함수 등록
        public XLCustomTemplate RegisterFunction(string name, FunctionHandler handler)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name));

            _functions[name.ToLowerInvariant()] = handler ??
                throw new ArgumentNullException(nameof(handler));

            return this;
        }

        // 확장 메서드: 글로벌 리졸버 설정
        public XLCustomTemplate SetGlobalResolver(GlobalVariableResolver resolver)
        {
            _interpreter.SetGlobalResolver(resolver);
            return this;
        }

        // IDisposable 구현
        public void Dispose()
        {
            if (_isDisposed) return;

            _workbook.Dispose();
            _isDisposed = true;
        }
    }
}