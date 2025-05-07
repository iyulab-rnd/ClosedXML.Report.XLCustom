namespace ClosedXML.Report.XLCustom.Internals
{
    /// <summary>
    /// StringBuilder extension methods
    /// </summary>
    internal static class StringBuilderExtensions
    {
        /// <summary>
        /// Appends a line with indentation
        /// </summary>
        public static StringBuilder AppendLineIndented(this StringBuilder sb, string text, int indentLevel = 0)
        {
            return sb.Append(new string(' ', indentLevel * 4)).AppendLine(text);
        }

        /// <summary>
        /// Appends a line with indentation and format
        /// </summary>
        public static StringBuilder AppendLineIndented(this StringBuilder sb, string format, int indentLevel = 0, params object[] args)
        {
            return sb.Append(new string(' ', indentLevel * 4)).AppendLine(string.Format(format, args));
        }
    }
}