namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Formatter registration and management for XLCustomTemplate
/// </summary>
public partial class XLCustomTemplate
{
    // Dictionary of registered formatters
    private readonly Dictionary<string, XLCustomFormatterFunc> _formatters = new Dictionary<string, XLCustomFormatterFunc>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Registers a formatter with the template
    /// </summary>
    public void RegisterFormat(string name, XLCustomFormatterFunc formatter)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (formatter == null)
            throw new ArgumentNullException(nameof(formatter));

        _formatters[name] = formatter;
    }

    /// <summary>
    /// Checks if a formatter is registered
    /// </summary>
    public bool HasFormatter(string name)
    {
        return !string.IsNullOrEmpty(name) && _formatters.ContainsKey(name);
    }

    /// <summary>
    /// Tries to get a formatter by name
    /// </summary>
    public bool TryGetFormatter(string name, out XLCustomFormatterFunc formatter)
    {
        if (string.IsNullOrEmpty(name))
        {
            formatter = null;
            return false;
        }

        return _formatters.TryGetValue(name, out formatter);
    }
}