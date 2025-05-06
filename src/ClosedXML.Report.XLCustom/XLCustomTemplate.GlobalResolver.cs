namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Global resolver functionality for XLCustomTemplate
    /// </summary>
    public partial class XLCustomTemplate
    {
        private Func<string, object> _globalResolver;

        /// <summary>
        /// Sets a global resolver function that will be used to resolve variables 
        /// that are not explicitly defined by AddVariable
        /// </summary>
        public void SetGlobalResolver(Func<string, object> resolver)
        {
            _globalResolver = resolver ?? throw new ArgumentNullException(nameof(resolver));
        }

        /// <summary>
        /// Attempts to resolve a variable using the global resolver
        /// </summary>
        internal bool TryResolveGlobal(string variableName, out object value)
        {
            if (_globalResolver != null)
            {
                try
                {
                    value = _globalResolver(variableName);
                    return value != null;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error in global resolver: {ex.Message}");
                    value = null;
                    return false;
                }
            }

            value = null;
            return false;
        }

        /// <summary>
        /// Resolves a variable directly, including through the global resolver
        /// </summary>
        public object Resolve(string variableName)
        {
            // Check variables first
            if (_variables.TryGetValue(variableName, out var value))
                return value;

            // Then try global resolver
            if (TryResolveGlobal(variableName, out value))
                return value;

            // Not found
            return null;
        }
    }
}