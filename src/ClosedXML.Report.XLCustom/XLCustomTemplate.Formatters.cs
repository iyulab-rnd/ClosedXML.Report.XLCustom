namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Formatter related functionality for XLCustomTemplate
    /// </summary>
    public partial class XLCustomTemplate
    {
        private readonly Dictionary<string, IXLCustomFormatter> _formatters = new();
        private readonly Dictionary<string, Func<object, string[], object>> _formatDelegates = new();

        /// <summary>
        /// Registers a custom formatter using a delegate
        /// </summary>
        public void RegisterFormat(string name, Func<object, string[], object> formatter)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Format name cannot be null or empty");

            if (formatter == null)
                throw new ArgumentNullException(nameof(formatter), "Formatter function cannot be null");

            _formatters.Remove(name);
            _formatDelegates.Remove(name);

            _formatDelegates[name] = formatter;
            Debug.WriteLine($"Registered format delegate: {name}");
        }

        /// <summary>
        /// Registers a custom formatter using an ICustomFormatter implementation
        /// </summary>
        public void RegisterFormat(string name, IXLCustomFormatter formatter)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Format name cannot be null or empty");

            if (formatter == null)
                throw new ArgumentNullException(nameof(formatter), "Formatter implementation cannot be null");

            _formatters.Remove(name);
            _formatDelegates.Remove(name);

            _formatters[name] = formatter;
            Debug.WriteLine($"Registered format implementation: {name}");
        }

        /// <summary>
        /// Gets a registered formatter by name
        /// </summary>
        public bool TryGetFormatter(string name, out Func<object, string[], object> formatter)
        {
            Debug.WriteLine($"Looking for formatter: {name}");

            if (_formatDelegates.TryGetValue(name, out var formatDelegate))
            {
                Debug.WriteLine($"Found formatter delegate: {name}");
                formatter = formatDelegate;
                return true;
            }

            if (_formatters.TryGetValue(name, out var customFormatter))
            {
                Debug.WriteLine($"Found formatter implementation: {name}");
                formatter = customFormatter.Format;
                return true;
            }

            Debug.WriteLine($"Formatter not found: {name}");
            formatter = null;
            return false;
        }

        /// <summary>
        /// Formats a value using a registered formatter
        /// </summary>
        public object FormatValue(string formatterName, object value, params string[] parameters)
        {
            if (string.IsNullOrEmpty(formatterName))
                throw new ArgumentNullException(nameof(formatterName));

            if (TryGetFormatter(formatterName, out var formatter))
            {
                return formatter(value, parameters ?? Array.Empty<string>());
            }

            throw new ArgumentException($"Formatter '{formatterName}' not found");
        }

        /// <summary>
        /// Lists all registered formatters
        /// </summary>
        public IEnumerable<string> GetRegisteredFormatters()
        {
            return _formatters.Keys.Concat(_formatDelegates.Keys).Distinct();
        }
    }
}