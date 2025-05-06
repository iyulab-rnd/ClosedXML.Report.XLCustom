namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Built-in formatters and functions registration for XLCustomTemplate
    /// </summary>
    public partial class XLCustomTemplate
    {
        /// <summary>
        /// Registers all built-in formatters and functions
        /// </summary>
        private void RegisterBuiltIns()
        {
            RegisterBuiltInFormatters();
            RegisterBuiltInFunctions();
        }

        /// <summary>
        /// Registers all built-in formatters
        /// </summary>
        private void RegisterBuiltInFormatters()
        {
            RegisterFormat("upper", BuiltInFormatters.Upper);
            RegisterFormat("lower", BuiltInFormatters.Lower);
            RegisterFormat("titlecase", BuiltInFormatters.TitleCase);
            RegisterFormat("mask", BuiltInFormatters.Mask);
            RegisterFormat("truncate", BuiltInFormatters.Truncate);
            RegisterFormat("currency", BuiltInFormatters.Currency);
            RegisterFormat("number", BuiltInFormatters.Number);
            RegisterFormat("percent", BuiltInFormatters.Percent);
            RegisterFormat("date", BuiltInFormatters.Date);
        }

        /// <summary>
        /// Registers all built-in functions
        /// </summary>
        private void RegisterBuiltInFunctions()
        {
            RegisterFunction("bold", BuiltInFunctions.Bold);
            RegisterFunction("italic", BuiltInFunctions.Italic);
            RegisterFunction("color", BuiltInFunctions.Color);
            RegisterFunction("link", BuiltInFunctions.Link);
            RegisterFunction("image", BuiltInFunctions.Image);
            RegisterFunction("format", BuiltInFunctions.Format);
            RegisterFunction("background", BuiltInFunctions.BackgroundColor);
            RegisterFunction("center", BuiltInFunctions.Center);
            RegisterFunction("border", BuiltInFunctions.Border);
        }
    }
}