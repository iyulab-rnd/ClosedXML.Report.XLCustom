namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Extended template class that provides advanced customization capabilities
/// while maintaining compatibility with ClosedXML.Report
/// </summary>
public partial class XLCustomTemplate : IXLTemplate
{
    private readonly XLTemplate _baseTemplate;
    private readonly bool _disposeWorkbookWithTemplate;
    private readonly TemplateErrors _errors = new();
    private readonly Dictionary<string, object> _variables = new();

    // Dictionary to store custom range interpreters by sheet
    private readonly Dictionary<string, CustomRangeInterpreter> _interpreters = new();
    private readonly List<string> _variablesAddedToBase = new();

    public bool IsDisposed { get; private set; }

    public IXLWorkbook Workbook { get; private set; }
    
    /// <summary>
    /// Initializes a new instance from a file
    /// </summary>
    public XLCustomTemplate(string fileName) : this(new XLWorkbook(fileName))
    {
        _disposeWorkbookWithTemplate = true;
    }

    /// <summary>
    /// Initializes a new instance from a stream
    /// </summary>
    public XLCustomTemplate(Stream stream) : this(new XLWorkbook(stream))
    {
        _disposeWorkbookWithTemplate = true;
    }

    /// <summary>
    /// Initializes a new instance from an existing workbook
    /// </summary>
    public XLCustomTemplate(IXLWorkbook workbook)
    {
        Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));

        // Create the base ClosedXML.Report template
        _baseTemplate = new XLTemplate(Workbook);

        // Initialize custom processors
        InitializeCustomProcessors();
    }

    /// <summary>
    /// Initializes custom tag processors
    /// </summary>
    private void InitializeCustomProcessors()
    {
        try
        {
            // Get access to RangeProcessor in base template
            var rangeProcessorField = typeof(XLTemplate).GetField("_rangeProcessor",
                BindingFlags.NonPublic | BindingFlags.Instance);

            if (rangeProcessorField != null)
            {
                var processor = rangeProcessorField.GetValue(_baseTemplate);

                if (processor != null)
                {
                    // Register custom handlers if needed
                    Debug.WriteLine("Custom processors initialized");
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to initialize custom processors: {ex.Message}");
            // Non-critical, continue with standard processing
        }
    }

    /// <summary>
    /// Gets or creates a custom interpreter for a specific worksheet and alias
    /// </summary>
    internal CustomRangeInterpreter GetInterpreter(string sheetName, string alias = null)
    {
        string key = $"{sheetName}:{alias ?? string.Empty}";

        if (!_interpreters.TryGetValue(key, out var interpreter))
        {
            interpreter = new CustomRangeInterpreter(alias, _errors, this);
            _interpreters[key] = interpreter;
        }

        return interpreter;
    }

    /// <summary>
    /// Creates a TemplateError with appropriate range
    /// </summary>
    private TemplateError CreateTemplateError(string message, IXLRange range = null)
    {
        // Create a dummy range if one is not provided
        if (range == null)
        {
            var ws = Workbook.Worksheet(1);
            if (ws == null && Workbook.Worksheets.Count > 0)
                ws = Workbook.Worksheets.First();

            if (ws != null)
                range = ws.Range("A1", "A1");
        }

        // If we still don't have a range, create a temporary one
        if (range == null)
        {
            var tempWorkbook = new XLWorkbook();
            var tempWorksheet = tempWorkbook.AddWorksheet("Temp");
            range = tempWorksheet.Range("A1", "A1");
        }

        return new TemplateError(message, range);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    public void Dispose()
    {
        if (IsDisposed)
            return;

        if (_disposeWorkbookWithTemplate)
        {
            Workbook.Dispose();
            _baseTemplate.Dispose();
        }

        Workbook = null;
        IsDisposed = true;
    }

    /// <summary>
    /// Checks if template is disposed
    /// </summary>
    private void CheckIsDisposed()
    {
        if (IsDisposed)
            throw new ObjectDisposedException("Template has been disposed");
    }
}