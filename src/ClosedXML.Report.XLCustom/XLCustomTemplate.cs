using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Report.Excel;
using ClosedXML.Report.Options;
using ClosedXML.Report.Utils;

namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Enhanced template engine that provides customization capabilities
    /// while maintaining compatibility with ClosedXML.Report
    /// </summary>
    public partial class XLCustomTemplate : IXLTemplate
    {
        private readonly bool _disposeWorkbookWithTemplate;
        private readonly TemplateErrors _errors = new TemplateErrors();
        private readonly FormulaEvaluator _evaluator;
        private readonly CustomFormulaEvaluator _customEvaluator;
        private RangeInterpreter _interpreter;

        private readonly Dictionary<string, object> _variables = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _addedVariables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool _interpreterInitialized = false;

        /// <summary>
        /// Gets the Excel workbook associated with this template
        /// </summary>
        public IXLWorkbook Workbook { get; private set; }

        /// <summary>
        /// Gets whether this template has been disposed
        /// </summary>
        public bool IsDisposed { get; private set; }

        /// <summary>
        /// Initializes a new instance from a file
        /// </summary>
        public XLCustomTemplate(string fileName)
            : this(new XLWorkbook(fileName))
        {
            _disposeWorkbookWithTemplate = true;
        }

        /// <summary>
        /// Initializes a new instance from a stream
        /// </summary>
        public XLCustomTemplate(Stream stream)
            : this(new XLWorkbook(stream))
        {
            _disposeWorkbookWithTemplate = true;
        }

        /// <summary>
        /// Initializes a new instance from an existing workbook
        /// </summary>
        public XLCustomTemplate(IXLWorkbook workbook)
        {
            // Register all the standard tags
            RegisterStandardTags();

            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));

            // Initialize evaluator
            _evaluator = new FormulaEvaluator();

            // Setup custom formula evaluator that integrates with the template
            _customEvaluator = new CustomFormulaEvaluator(this, _evaluator);

            // Initialize interpreter without adding variables yet
            _interpreter = new RangeInterpreter(null, _errors);

            // Register custom resolver for the evaluator through the evaluator
            _evaluator.SetGlobalResolver(variableName => ResolveVariable(variableName));

            // Register built-in formatters and functions
            RegisterBuiltIns();
        }

        /// <summary>
        /// Generates the report by processing all template expressions
        /// </summary>
        public XLGenerateResult Generate()
        {
            CheckIsDisposed();

            try
            {
                // Process collection metadata expressions
                ProcessCollectionMetadata();

                // Preprocess enhanced expressions to make them compatible with base engine
                PreprocessEnhancedExpressions();

                // 모든 변수(임시 변수 포함)가 추가된 후 인터프리터 초기화
                EnsureInterpreterInitialized();

                // Process each visible worksheet
                var worksheets = Workbook.Worksheets.Where(sh =>
                    sh.Visibility == XLWorksheetVisibility.Visible &&
                    !sh.PivotTables.Any()).ToList();

                foreach (var worksheet in worksheets)
                {
                    worksheet.ReplaceCFFormulaeToR1C1();
                    _interpreter.Evaluate(worksheet.AsRange());
                    worksheet.ReplaceCFFormulaeToA1();
                }

                // Post-process any remaining enhanced expressions
                PostprocessEnhancedExpressions();

                return new XLGenerateResult(_errors);
            }
            catch (Exception ex)
            {
                _errors.Add(new TemplateError($"Unexpected error: {ex.Message}", null));
                return new XLGenerateResult(_errors);
            }
        }

        /// <summary>
        /// Adds variables from an object
        /// </summary>
        public void AddVariable(object value)
        {
            CheckIsDisposed();

            if (value is IDictionary dictionary)
            {
                foreach (DictionaryEntry entry in dictionary)
                {
                    AddVariable(entry.Key.ToString(), entry.Value);
                }
            }
            else
            {
                var type = value.GetType();
                var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance)
                    .Where(f => f.IsPublic)
                    .Select(f => new { f.Name, val = f.GetValue(value), type = f.FieldType })
                    .Concat(
                        type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .Where(f => f.CanRead)
                            .Select(f => new
                            {
                                f.Name,
                                val = f.GetValue(value, Array.Empty<object>()),
                                type = f.PropertyType
                            })
                    );

                foreach (var field in fields)
                {
                    AddVariable(field.Name, field.val);
                }
            }
        }

        /// <summary>
        /// Adds a variable with a specific name
        /// </summary>
        public void AddVariable(string alias, object value)
        {
            CheckIsDisposed();

            // For primitive types and strings, store as-is without any conversion
            if (value == null || value.GetType().IsPrimitive() || value is string)
            {
                // Store variable in our internal dictionary (overwrites if exists)
                _variables[alias] = value;

                // Add to evaluator (overwrites if exists)
                _evaluator.AddVariable(alias, value);

                // Mark interpreter as not initialized so variables will be updated
                _interpreterInitialized = false;

                return;
            }

            // For other types, store in our variables dictionary
            _variables[alias] = value;

            // Add to evaluator
            _evaluator.AddVariable(alias, value);

            // Mark interpreter as not initialized so variables will be updated
            _interpreterInitialized = false;
        }

        /// <summary>
        /// Ensures that all variables are properly initialized in the RangeInterpreter
        /// </summary>
        private void EnsureInterpreterInitialized()
        {
            if (_interpreterInitialized)
                return;

            // Create a new interpreter
            _interpreter = new RangeInterpreter(null, _errors);
            _addedVariables.Clear();

            // Add all current variables
            foreach (var kvp in _variables)
            {
                if (kvp.Value == null)
                {
                    // Null values can be added directly
                    _interpreter.AddVariable(kvp.Key, null);
                    _addedVariables.Add(kvp.Key);
                    continue;
                }

                // Handle DataTable conversion for collections
                object baseValue = kvp.Value;
                if (baseValue is DataTable dt)
                {
                    baseValue = dt.Rows.Cast<DataRow>();
                }
                else if (baseValue is IEnumerable enumerable && !(baseValue is string))
                {
                    // Attempt to convert non-generic IEnumerable to a typed List
                    var itemType = enumerable.GetItemType();
                    if (itemType != typeof(object))
                    {
                        try
                        {
                            baseValue = ConvertToTypedList(enumerable, itemType);
                        }
                        catch
                        {
                            // If conversion fails, keep the original value
                            baseValue = kvp.Value;
                        }
                    }
                }

                try
                {
                    // Add to interpreter
                    _interpreter.AddVariable(kvp.Key, baseValue);
                    _addedVariables.Add(kvp.Key);
                }
                catch (Exception ex)
                {
                    // Log error but continue with other variables
                    _errors.Add(new TemplateError(
                        $"Error adding variable '{kvp.Key}': {ex.Message}", null));
                }
            }

            _interpreterInitialized = true;
        }

        /// <summary>
        /// Checks if template is disposed
        /// </summary>
        private void CheckIsDisposed()
        {
            if (IsDisposed)
                throw new ObjectDisposedException("Template has been disposed");
        }

        /// <summary>
        /// Disposes of resources
        /// </summary>
        public void Dispose()
        {
            if (IsDisposed)
                return;

            if (_disposeWorkbookWithTemplate)
            {
                Workbook?.Dispose();
            }

            Workbook = null;
            IsDisposed = true;
        }

        /// <summary>
        /// Converts an IEnumerable to a typed list
        /// </summary>
        private static IEnumerable ConvertToTypedList(IEnumerable enumerable, Type itemType)
        {
            var listType = typeof(List<>).MakeGenericType(itemType);
            var list = (IList)Activator.CreateInstance(listType);

            foreach (var item in enumerable)
            {
                list.Add(item);
            }

            return list;
        }
    }
}