namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Function registration and management for XLCustomTemplate
/// </summary>
public partial class XLCustomTemplate
{
    // Dictionary of registered functions
    private readonly Dictionary<string, XLCustomFunctionFunc> _functions = new Dictionary<string, XLCustomFunctionFunc>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Registers a function with the template
    /// </summary>
    public void RegisterFunction(string name, XLCustomFunctionFunc function)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (function == null)
            throw new ArgumentNullException(nameof(function));

        _functions[name] = function;
    }

    /// <summary>
    /// Checks if a function is registered
    /// </summary>
    public bool HasFunction(string name)
    {
        return !string.IsNullOrEmpty(name) && _functions.ContainsKey(name);
    }

    /// <summary>
    /// Tries to get a function by name
    /// </summary>
    public bool TryGetFunction(string name, out XLCustomFunctionFunc function)
    {
        if (string.IsNullOrEmpty(name))
        {
            function = null;
            return false;
        }

        return _functions.TryGetValue(name, out function);
    }
}