using System;
using System.Globalization;
using System.Text;

namespace ClosedXML.Report.XLCustom.Formatters;

/// <summary>
/// Provides built-in formatters that can be registered with XLCustomTemplate
/// </summary>
public static class BuiltInFormatters
{
    /// <summary>
    /// Formats text as uppercase
    /// </summary>
    public static readonly IXLCustomFormatter Upper = new DelegateFormatter(
        (value, parameters) => value?.ToString()?.ToUpper()
    );

    /// <summary>
    /// Formats text as lowercase
    /// </summary>
    public static readonly IXLCustomFormatter Lower = new DelegateFormatter(
        (value, parameters) => value?.ToString()?.ToLower()
    );

    /// <summary>
    /// Formats text as title case
    /// </summary>
    public static readonly IXLCustomFormatter TitleCase = new DelegateFormatter(
        (value, parameters) =>
        {
            if (value == null) return null;
            var text = value.ToString();
            var textInfo = CultureInfo.CurrentCulture.TextInfo;
            return textInfo.ToTitleCase(text.ToLower());
        }
    );

    /// <summary>
    /// Formats a value by applying a mask
    /// </summary>
    public static readonly IXLCustomFormatter Mask = new DelegateFormatter(
        (value, parameters) =>
        {
            if (value == null || parameters.Length == 0) return value;

            string text = value.ToString();
            string mask = parameters[0];
            int textIndex = 0;
            var result = new StringBuilder();

            foreach (char c in mask)
            {
                if (c == '#')
                {
                    if (textIndex < text.Length)
                        result.Append(text[textIndex++]);
                }
                else
                {
                    result.Append(c);
                }
            }

            return result.ToString();
        }
    );

    /// <summary>
    /// Formats a value as a phone number
    /// </summary>
    public static readonly IXLCustomFormatter Phone = new DelegateFormatter(
        (value, parameters) =>
        {
            if (value == null) return null;

            string text = value.ToString().Trim().Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "");
            string format = parameters.Length > 0 ? parameters[0] : "(###) ###-####";

            return Mask.Format(text, new[] { format });
        }
    );

    /// <summary>
    /// Truncates text to specified length
    /// </summary>
    public static readonly IXLCustomFormatter Truncate = new DelegateFormatter(
        (value, parameters) =>
        {
            if (value == null) return null;

            string text = value.ToString();
            if (parameters.Length == 0 || !int.TryParse(parameters[0], out int length))
                return text;

            if (text.Length <= length)
                return text;

            string suffix = parameters.Length > 1 ? parameters[1] : "...";
            return text.Substring(0, length) + suffix;
        }
    );

    /// <summary>
    /// A formatter implementation that uses a delegate function
    /// </summary>
    private class DelegateFormatter : IXLCustomFormatter
    {
        private readonly Func<object, string[], object> _formatFunction;

        public DelegateFormatter(Func<object, string[], object> formatFunction)
        {
            _formatFunction = formatFunction ?? throw new ArgumentNullException(nameof(formatFunction));
        }

        public object Format(object value, string[] parameters)
        {
            return _formatFunction(value, parameters);
        }
    }
}