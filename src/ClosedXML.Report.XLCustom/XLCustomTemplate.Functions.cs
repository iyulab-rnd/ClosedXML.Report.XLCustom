using static System.Runtime.InteropServices.JavaScript.JSType;

namespace ClosedXML.Report.XLCustom
{
    /// <summary>
    /// Function related functionality for XLCustomTemplate
    /// </summary>
    public partial class XLCustomTemplate
    {
        private readonly Dictionary<string, IXLCustomFunction> _functions = new();
        private readonly Dictionary<string, Action<IXLCell, object, string[]>> _functionDelegates = new();

        /// <summary>
        /// Registers a custom function using a delegate
        /// </summary>
        public void RegisterFunction(string name, Action<IXLCell, object, string[]> function)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Function name cannot be null or empty");

            if (function == null)
                throw new ArgumentNullException(nameof(function), "Function delegate cannot be null");

            _functions.Remove(name);
            _functionDelegates.Remove(name);

            _functionDelegates[name] = function;
            Debug.WriteLine($"Registered function delegate: {name}");
        }

        /// <summary>
        /// Registers a custom function using an ICustomFunction implementation
        /// </summary>
        public void RegisterFunction(string name, IXLCustomFunction function)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentNullException(nameof(name), "Function name cannot be null or empty");

            if (function == null)
                throw new ArgumentNullException(nameof(function), "Function implementation cannot be null");

            _functions.Remove(name);
            _functionDelegates.Remove(name);

            _functions[name] = function;
            Debug.WriteLine($"Registered function implementation: {name}");
        }

        /// <summary>
        /// Gets a registered function by name
        /// </summary>
        public bool TryGetFunction(string name, out Action<IXLCell, object, string[]> function)
        {
            Debug.WriteLine($"Looking for function with name: '{name}'");

            if (_functionDelegates.TryGetValue(name, out var functionDelegate))
            {
                Debug.WriteLine($"Found function delegate for name: '{name}'");
                function = functionDelegate;
                return true;
            }

            if (_functions.TryGetValue(name, out var customFunction))
            {
                Debug.WriteLine($"Found custom function for name: '{name}'");
                function = customFunction.Process;
                return true;
            }

            Debug.WriteLine($"Function not found: '{name}'");
            function = null;
            return false;
        }

        /// <summary>
        /// Directly applies a function to a cell
        /// </summary>
        public void ApplyFunction(IXLCell cell, string functionName, object value, string[] parameters = null)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));

            if (string.IsNullOrEmpty(functionName))
                throw new ArgumentNullException(nameof(functionName));

            Debug.WriteLine($"Applying function {functionName} to cell {cell.Address}");

            if (!TryGetFunction(functionName, out var function))
            {
                Debug.WriteLine($"Function {functionName} not found");
                throw new ArgumentException($"Function '{functionName}' not found", nameof(functionName));
            }

            Debug.WriteLine($"Applying function with value: {value}");
            function(cell, value, parameters ?? Array.Empty<string>());
            Debug.WriteLine($"Function applied, cell value: {cell.Value}, Bold: {cell.Style.Font.Bold}");
        }

        /// <summary>
        /// Lists all registered functions
        /// </summary>
        public IEnumerable<string> GetRegisteredFunctions()
        {
            return _functions.Keys.Concat(_functionDelegates.Keys).Distinct();
        }

        /// <summary>
        /// Debug helper to output information about registered functions
        /// </summary>
        public void DebugFunctions()
        {
            Debug.WriteLine("\n==== Registered Functions ====");

            foreach (var func in _functions)
            {
                Debug.WriteLine($"Function: {func.Key} (IXLCustomFunction)");
            }

            foreach (var func in _functionDelegates)
            {
                Debug.WriteLine($"Function: {func.Key} (Delegate)");
            }

            Debug.WriteLine("==== Built-in Functions ====");
            if (TryGetFunction("bold", out var boldFunc))
            {
                Debug.WriteLine("Bold function is registered and available");
            }
            else
            {
                Debug.WriteLine("Bold function is NOT registered");
            }

            Debug.WriteLine("=============================\n");
        }

        /// <summary>
        /// Tests a function on a cell for debugging
        /// </summary>
        public void TestFunction(string functionName, IXLCell testCell, object testValue)
        {
            if (TryGetFunction(functionName, out var function))
            {
                Debug.WriteLine($"Testing function {functionName} on cell {testCell.Address}");

                var originalValue = testCell.Value;
                var originalBold = testCell.Style.Font.Bold;

                try
                {
                    function(testCell, testValue, Array.Empty<string>());
                    Debug.WriteLine($"Function applied. Cell value: {testCell.Value}, Bold: {testCell.Style.Font.Bold}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Function test failed: {ex.Message}");

                    // Restore cell
                    testCell.Value = originalValue;
                    testCell.Style.Font.Bold = originalBold;
                }
            }
            else
            {
                Debug.WriteLine($"Function {functionName} not found for testing");
            }
        }
    }
}