namespace ClosedXML.Report.XLCustom.Internals
{
    /// <summary>
    /// Helper extension methods
    /// </summary>
    internal static class Extensions
    {
        /// <summary>
        /// Checks if a string contains format expression syntax
        /// </summary>
        public static bool ContainsFormatExpression(this string source)
        {
            if (string.IsNullOrEmpty(source))
                return false;

            int startIndex = source.IndexOf("{{");
            if (startIndex < 0) return false;

            int endIndex = source.IndexOf("}}", startIndex);
            if (endIndex < 0) return false;

            int formatIndex = source.IndexOf(":", startIndex);
            return formatIndex > startIndex && formatIndex < endIndex;
        }

        /// <summary>
        /// Checks if a string contains function expression syntax
        /// </summary>
        public static bool ContainsFunctionExpression(this string source)
        {
            if (string.IsNullOrEmpty(source))
                return false;

            int startIndex = source.IndexOf("{{");
            if (startIndex < 0) return false;

            int endIndex = source.IndexOf("}}", startIndex);
            if (endIndex < 0) return false;

            int pipeIndex = source.IndexOf("|", startIndex);
            return pipeIndex > startIndex && pipeIndex < endIndex;
        }
    }
}