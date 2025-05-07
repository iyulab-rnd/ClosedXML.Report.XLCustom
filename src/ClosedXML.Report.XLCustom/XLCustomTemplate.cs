using ClosedXML.Excel;
using ClosedXML.Report;
using ClosedXML.Report.Options;
using ClosedXML.Report.XLCustom.Tags;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Extends the XLTemplate class with enhanced expression handling capabilities
    /// </summary>
    public partial class XLCustomTemplate : IXLTemplate
    {
        private XLTemplate _baseTemplate;
        private readonly FormatRegistry _formatRegistry;
        private readonly FunctionRegistry _functionRegistry;
        private readonly IXLWorkbook _workbook;
        private bool _preprocessed = false;
        private bool _templateCreated = false;

        static XLCustomTemplate()
        {
            // 태그 등록 - 대소문자 주의
            try
            {
                // 우리 커스텀 태그들 등록 (태그 이름이 customformat, customfunction 임에 주의)
                TagsRegister.Add<CustomFormatTag>("customformat", 150);
                TagsRegister.Add<CustomFunctionTag>("customfunction", 140);
                TagsRegister.Add<FormatTag>("format", 120);
            }
            catch (Exception ex)
            {
                Log.Debug($"Tag registration error: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a new instance of XLCustomTemplate from file path
        /// </summary>
        public XLCustomTemplate(string filePath)
        {
            _formatRegistry = new FormatRegistry();
            _functionRegistry = new FunctionRegistry();

            // 레지스트리를 싱글톤에 설정하여 태그들이 접근할 수 있게 함
            XLCustomRegistry.Instance.SetRegistries(_formatRegistry, _functionRegistry);

            // 워크북만 로드하고, 전처리는 나중에 수행
            _workbook = new XLWorkbook(filePath);
        }

        /// <summary>
        /// Creates a new instance of XLCustomTemplate from a stream
        /// </summary>
        public XLCustomTemplate(Stream stream)
        {
            _formatRegistry = new FormatRegistry();
            _functionRegistry = new FunctionRegistry();

            // 레지스트리를 싱글톤에 설정하여 태그들이 접근할 수 있게 함
            XLCustomRegistry.Instance.SetRegistries(_formatRegistry, _functionRegistry);

            // 워크북만 로드하고, 전처리는 나중에 수행
            _workbook = new XLWorkbook(stream);
        }

        /// <summary>
        /// Creates a new instance of XLCustomTemplate from an existing workbook
        /// </summary>
        public XLCustomTemplate(IXLWorkbook workbook)
        {
            _formatRegistry = new FormatRegistry();
            _functionRegistry = new FunctionRegistry();

            // 레지스트리를 싱글톤에 설정하여 태그들이 접근할 수 있게 함
            XLCustomRegistry.Instance.SetRegistries(_formatRegistry, _functionRegistry);

            // 워크북 참조만 저장하고, 전처리는 나중에 수행
            _workbook = workbook;
        }

        /// <summary>
        /// Gets the underlying workbook
        /// </summary>
        public IXLWorkbook Workbook => _templateCreated ? _baseTemplate.Workbook : _workbook;

        /// <summary>
        /// Registers a custom format processor
        /// </summary>
        public XLCustomTemplate RegisterFormat(string formatName, XLFormatHandler formatter)
        {
            Log.Debug($"Registering format: {formatName}");
            _formatRegistry.Register(formatName, formatter);

            // 싱글톤 레지스트리에 최신 포맷 레지스트리 설정
            XLCustomRegistry.Instance.SetRegistries(_formatRegistry, _functionRegistry);

            _preprocessed = false; // 포맷 등록 시 재전처리 필요
            return this;
        }

        /// <summary>
        /// Registers a custom function processor that can manipulate cells
        /// </summary>
        public XLCustomTemplate RegisterFunction(string functionName, XLFunctionHandler function)
        {
            _functionRegistry.Register(functionName, function);
            _preprocessed = false; // 함수 등록 시 재전처리 필요
            return this;
        }

        /// <summary>
        /// Registers all built-in formatters and functions
        /// </summary>
        public XLCustomTemplate RegisterBuiltIns()
        {
            RegisterBuiltInFormatters();
            RegisterBuiltInFunctions();
            return this;
        }

        /// <summary>
        /// Registers built-in formatters
        /// </summary>
        public XLCustomTemplate RegisterBuiltInFormatters()
        {
            RegisterFormat("upper", (value, _) => value?.ToString()?.ToUpper());
            RegisterFormat("lower", (value, _) => value?.ToString()?.ToLower());
            RegisterFormat("titlecase", (value, _) => System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value?.ToString() ?? ""));
            RegisterFormat("mask", (value, parameters) => {
                var mask = parameters.Length > 0 ? parameters[0] : "";
                return ApplyMask(value?.ToString() ?? "", mask);
            });
            RegisterFormat("truncate", (value, parameters) => {
                var text = value?.ToString() ?? "";
                var maxLength = parameters.Length > 0 && int.TryParse(parameters[0], out var len) ? len : 50;
                var suffix = parameters.Length > 1 ? parameters[1] : "...";
                return text.Length > maxLength ? text.Substring(0, maxLength) + suffix : text;
            });
            RegisterFormat("phone", (value, parameters) => {
                if (value == null) return null;
                string text = value.ToString();
                if (string.IsNullOrEmpty(text)) return text;

                if (parameters.Length == 0) return text;

                string format = parameters[0];
                return ApplyMask(text, format);
            });
            RegisterFormat("currency", (value, parameters) => {
                if (value == null) return null;

                if (decimal.TryParse(value.ToString(), out decimal amount))
                {
                    string currencySymbol = parameters.Length > 0 ? parameters[0] : "$";
                    return $"{currencySymbol}{amount:N2}";
                }

                return value;
            });

            return this;
        }

        /// <summary>
        /// Registers built-in functions
        /// </summary>
        public XLCustomTemplate RegisterBuiltInFunctions()
        {
            RegisterFunction("bold", (cell, value, _) => {
                cell.SetValue(value);
                cell.Style.Font.Bold = true;
            });

            RegisterFunction("italic", (cell, value, _) => {
                cell.SetValue(value);
                cell.Style.Font.Italic = true;
            });

            RegisterFunction("color", (cell, value, parameters) => {
                cell.SetValue(value);
                var colorName = parameters.Length > 0 ? parameters[0] : "Black";
                try
                {
                    cell.Style.Font.FontColor = XLColor.FromName(colorName);
                }
                catch (Exception ex)
                {
                    Log.Debug($"Error setting color: {ex.Message}");
                    cell.Style.Font.FontColor = XLColor.Black;
                }
            });

            RegisterFunction("link", (cell, value, parameters) => {
                string url = value?.ToString();
                if (string.IsNullOrEmpty(url)) return;

                string text = parameters.Length > 0 ? parameters[0] : url;
                cell.Value = text;
                cell.SetHyperlink(new XLHyperlink(url));
            });

            RegisterFunction("image", (cell, value, parameters) => {
                string imagePath = value?.ToString();
                if (string.IsNullOrEmpty(imagePath)) return;

                try
                {
                    var picture = cell.Worksheet.AddPicture(imagePath);

                    // Apply optional size parameters
                    if (parameters.Length > 0 && int.TryParse(parameters[0], out int width))
                        picture.Width = width;

                    if (parameters.Length > 1 && int.TryParse(parameters[1], out int height))
                        picture.Height = height;

                    picture.MoveTo(cell);
                }
                catch (Exception ex)
                {
                    Log.Debug($"Error adding image: {ex.Message}");
                    cell.Value = $"Error: {ex.Message}";
                    cell.Style.Font.FontColor = XLColor.Red;
                }
            });

            return this;
        }

        /// <summary>
        /// Force workbook preprocessing
        /// </summary>
        public XLCustomTemplate Preprocess()
        {
            EnsurePreprocessed();
            return this;
        }

        // IXLTemplate 인터페이스 구현
        public XLGenerateResult Generate()
        {
            EnsureTemplateCreated();
            return _baseTemplate.Generate();
        }

        public void AddVariable(object value)
        {
            EnsureTemplateCreated();
            _baseTemplate.AddVariable(value);
        }

        public void AddVariable(string alias, object value)
        {
            EnsureTemplateCreated();
            _baseTemplate.AddVariable(alias, value);
        }

        public void Dispose()
        {
            if (_templateCreated)
            {
                _baseTemplate.Dispose();
            }
            else
            {
                _workbook.Dispose();
            }
        }

        private void EnsurePreprocessed()
        {
            if (!_preprocessed)
            {
                // ExpressionProcessor를 여기서 생성
                var expressionProcessor = new XLExpressionProcessor(_formatRegistry, _functionRegistry);
                PreprocessWorkbook(_workbook, expressionProcessor);
                _preprocessed = true;
            }
        }

        private void EnsureTemplateCreated()
        {
            if (!_templateCreated)
            {
                EnsurePreprocessed(); // 전처리 확인

                _baseTemplate = new XLTemplate(_workbook);
                _templateCreated = true;

                Log.Debug("Base XLTemplate created from processed workbook");
            }
        }

        private void PreprocessWorkbook(IXLWorkbook workbook, XLExpressionProcessor expressionProcessor)
        {
            Log.Debug("Starting workbook preprocessing...");

            // 워크북의 모든 시트를 처리
            foreach (var worksheet in workbook.Worksheets)
            {
                Log.Debug($"Processing worksheet: {worksheet.Name}");

                // 사용된 모든 셀 중 커스텀 표현식을 포함할 수 있는 셀 처리
                foreach (var cell in worksheet.CellsUsed())
                {
                    if (cell.HasFormula) continue; // 수식 셀은 건너뛰기

                    if (cell.DataType == XLDataType.Text)
                    {
                        var value = cell.GetString();
                        if (string.IsNullOrEmpty(value)) continue;

                        // 셀 내용 처리 및 필요한 경우 호환되는 태그로 대체
                        var newValue = expressionProcessor.ProcessExpression(value, cell);
                        if (newValue != value)
                        {
                            Log.Debug($"Cell {cell.Address}: Replaced '{value}' with '{newValue}'");
                            cell.Value = newValue;
                        }
                    }
                }
            }

            Log.Debug("Workbook preprocessing completed");
        }

        private string ApplyMask(string input, string mask)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(mask))
                return input;

            var result = new StringBuilder();
            var index = 0;

            foreach (var c in mask)
            {
                if (c == '#')
                {
                    if (index < input.Length)
                    {
                        result.Append(input[index]);
                        index++;
                    }
                }
                else
                    result.Append(c);
            }

            return result.ToString();
        }

        /// <summary>
        /// 디버깅용 메서드: 표현식을 처리하고 결과 반환
        /// </summary>
        public string DebugExpression(string expression, IXLCell cell = null)
        {
            var expressionProcessor = new XLExpressionProcessor(_formatRegistry, _functionRegistry);
            return expressionProcessor.ProcessExpression(expression, cell);
        }
    }
}