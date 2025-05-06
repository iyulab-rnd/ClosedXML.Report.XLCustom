namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Provides built-in formatters that can be registered with XLCustomTemplate
/// </summary>
public static class BuiltInFormatters
{
    /// <summary>
    /// Formats text as uppercase
    /// </summary>
    public static readonly XLCustomFormatterFunc Upper = (value, parameters) =>
        value?.ToString()?.ToUpper();

    /// <summary>
    /// Formats text as lowercase
    /// </summary>
    public static readonly XLCustomFormatterFunc Lower = (value, parameters) =>
        value?.ToString()?.ToLower();

    /// <summary>
    /// Formats text as title case
    /// </summary>
    public static readonly XLCustomFormatterFunc TitleCase = (value, parameters) =>
    {
        if (value == null) return null;
        var text = value.ToString();
        var textInfo = CultureInfo.CurrentCulture.TextInfo;
        return textInfo.ToTitleCase(text.ToLower());
    };

    /// <summary>
    /// Formats a value by applying a mask
    /// </summary>
    public static readonly XLCustomFormatterFunc Mask = (value, parameters) =>
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
    };

    /// <summary>
    /// Truncates text to specified length
    /// </summary>
    public static readonly XLCustomFormatterFunc Truncate = (value, parameters) =>
    {
        if (value == null) return null;

        string text = value.ToString();
        if (parameters.Length == 0 || !int.TryParse(parameters[0], out int length))
            return text;

        if (text.Length <= length)
            return text;

        // 기본 접미사는 "..."
        string suffix = parameters.Length > 1 ? parameters[1] : "...";

        // 텍스트를 자른 후 접미사 추가
        return text.Substring(0, length) + suffix;
    };

    /// <summary>
    /// Formats a currency value
    /// </summary>
    public static readonly XLCustomFormatterFunc Currency = (value, parameters) =>
    {
        if (value == null) return null;

        if (!decimal.TryParse(value.ToString(), out decimal amount))
            return value;

        string currencyCode = parameters.Length > 0 ? parameters[0] : "USD";

        try
        {
            var culture = GetCultureForCurrency(currencyCode);
            return amount.ToString("C", culture);
        }
        catch
        {
            // Fallback to current culture
            return amount.ToString("C");
        }
    };

    /// <summary>
    /// Formats a number with thousands separators
    /// </summary>
    public static readonly XLCustomFormatterFunc Number = (value, parameters) =>
    {
        if (value == null) return null;

        if (!decimal.TryParse(value.ToString(), out decimal number))
            return value;

        int decimals = parameters.Length > 0 && int.TryParse(parameters[0], out int d) ? d : 0;
        return number.ToString($"N{decimals}");
    };

    /// <summary>
    /// Formats a value as a percentage
    /// </summary>
    public static readonly XLCustomFormatterFunc Percent = (value, parameters) =>
    {
        if (value == null) return null;

        if (!decimal.TryParse(value.ToString(), out decimal number))
            return value;

        int decimals = parameters.Length > 0 && int.TryParse(parameters[0], out int d) ? d : 0;
        return number.ToString($"P{decimals}");
    };

    /// <summary>
    /// Formats a DateTime value
    /// </summary>
    public static readonly XLCustomFormatterFunc Date = (value, parameters) =>
    {
        if (value == null) return null;

        if (!(value is DateTime date) && !DateTime.TryParse(value.ToString(), out date))
            return value;

        string format = parameters.Length > 0 ? parameters[0] : "d";
        return date.ToString(format);
    };

    /// <summary>
    /// Gets a culture info for a currency code
    /// </summary>
    private static CultureInfo GetCultureForCurrency(string currencyCode)
    {
        return currencyCode.ToUpper() switch
        {
            "USD" => new CultureInfo("en-US"),
            "EUR" => new CultureInfo("fr-FR"),
            "GBP" => new CultureInfo("en-GB"),
            "JPY" => new CultureInfo("ja-JP"),
            "CAD" => new CultureInfo("en-CA"),
            "AUD" => new CultureInfo("en-AU"),
            _ => CultureInfo.CurrentCulture
        };
    }
}