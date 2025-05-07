using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Registry for custom format handlers
    /// </summary>
    public class FormatRegistry
    {
        private readonly Dictionary<string, XLFormatHandler> _formatters =
            new Dictionary<string, XLFormatHandler>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers a custom formatter
        /// </summary>
        public void Register(string formatName, XLFormatHandler formatter)
        {
            _formatters[formatName] = formatter ?? throw new ArgumentNullException(nameof(formatter));
        }

        /// <summary>
        /// Checks if a format is registered
        /// </summary>
        public bool IsRegistered(string formatName)
        {
            return _formatters.ContainsKey(formatName);
        }

        /// <summary>
        /// Applies a formatter to a value
        /// </summary>
        public object ApplyFormat(string formatName, object value, string[] parameters)
        {
            if (!IsRegistered(formatName))
                return value;

            return _formatters[formatName](value, parameters);
        }
    }

    /// <summary>
    /// Registry for custom function handlers
    /// </summary>
    public class FunctionRegistry
    {
        private readonly Dictionary<string, XLFunctionHandler> _functions =
            new Dictionary<string, XLFunctionHandler>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers a custom function
        /// </summary>
        public void Register(string functionName, XLFunctionHandler function)
        {
            _functions[functionName] = function ?? throw new ArgumentNullException(nameof(function));
        }

        /// <summary>
        /// Checks if a function is registered
        /// </summary>
        public bool IsRegistered(string functionName)
        {
            return _functions.ContainsKey(functionName);
        }

        /// <summary>
        /// Gets a registered function by name
        /// </summary>
        public XLFunctionHandler GetFunction(string functionName)
        {
            if (!IsRegistered(functionName))
                return null;

            return _functions[functionName];
        }
    }
}