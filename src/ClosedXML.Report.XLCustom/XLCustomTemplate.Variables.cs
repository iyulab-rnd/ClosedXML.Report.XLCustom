namespace ClosedXML.Report.XLCustom
{
    public partial class XLCustomTemplate
    {
        /// <summary>
        /// Sets a global resolver function that will be used to resolve variables 
        /// that are not explicitly defined by AddVariable
        /// </summary>
        public void SetGlobalResolver(Func<string, object> resolver)
        {
            _evaluator.SetGlobalResolver(resolver);
        }

        /// <summary>
        /// Resolves a variable directly
        /// </summary>
        internal object ResolveVariable(string variableName)
        {
            // Check internal variables dictionary
            if (_variables.TryGetValue(variableName, out var value))
            {
                return value;
            }

            // Try evaluator's global resolver
            if (_evaluator.TryResolveGlobal(variableName, out var resolvedValue))
            {
                return resolvedValue;
            }

            // Check collection counts
            if (variableName.Contains("_Count"))
            {
                string collectionName = variableName.Substring(0, variableName.Length - 6);
                if (TryGetCollection(collectionName, out var collection))
                {
                    return collection.Count;
                }
            }

            // Check nested properties
            if (variableName.Contains("."))
            {
                string[] parts = variableName.Split(new[] { '.' }, 2);
                string objectName = parts[0];
                string propertyPath = parts[1];

                if (_variables.TryGetValue(objectName, out var parentObj) && parentObj != null)
                {
                    try
                    {
                        return ResolveNestedProperty(parentObj, propertyPath);
                    }
                    catch
                    {
                        // If property resolution fails, continue with other checks
                    }
                }
            }

            // Check array indexers
            if (variableName.Contains("[") && variableName.Contains("]"))
            {
                try
                {
                    return ResolveArrayIndexer(variableName);
                }
                catch
                {
                    // If indexer resolution fails, continue with other checks
                }
            }

            return null;
        }

        /// <summary>
        /// Resolves a nested property path on an object
        /// </summary>
        private object ResolveNestedProperty(object obj, string propertyPath)
        {
            if (obj == null || string.IsNullOrEmpty(propertyPath))
                return null;

            // Handle array indexers in property path
            if (propertyPath.Contains("[") && propertyPath.Contains("]"))
            {
                // Split at the first dot or first square bracket, whichever comes first
                int dotIndex = propertyPath.IndexOf('.');
                int bracketIndex = propertyPath.IndexOf('[');

                if (dotIndex == -1 || (bracketIndex != -1 && bracketIndex < dotIndex))
                {
                    // We have an array indexer first
                    string arrayPath = bracketIndex == 0 ? "" : propertyPath.Substring(0, bracketIndex);

                    if (string.IsNullOrEmpty(arrayPath))
                    {
                        // Direct indexer on the object itself
                        // First, find the closing bracket
                        int closeBracket = propertyPath.IndexOf(']');
                        if (closeBracket == -1)
                            throw new FormatException("Invalid array indexer format: missing closing bracket");

                        string indexStr = propertyPath.Substring(bracketIndex + 1, closeBracket - bracketIndex - 1);
                        if (!int.TryParse(indexStr, out int index))
                            throw new FormatException($"Invalid array index: {indexStr}");

                        if (obj is IList list)
                        {
                            if (index < 0 || index >= list.Count)
                                return null; // Index out of range

                            obj = list[index];

                            // Check if there's more to the path
                            if (closeBracket + 1 < propertyPath.Length)
                            {
                                if (propertyPath[closeBracket + 1] == '.')
                                {
                                    // Continue with the rest of the path
                                    return ResolveNestedProperty(obj, propertyPath.Substring(closeBracket + 2));
                                }
                                else if (propertyPath[closeBracket + 1] == '[')
                                {
                                    // Another indexer
                                    return ResolveNestedProperty(obj, propertyPath.Substring(closeBracket + 1));
                                }
                            }

                            return obj;
                        }
                        else
                        {
                            throw new InvalidOperationException($"Object is not a list or array: {obj.GetType().Name}");
                        }
                    }
                    else
                    {
                        // Property followed by indexer
                        obj = GetPropertyValue(obj, arrayPath);
                        return ResolveNestedProperty(obj, propertyPath.Substring(bracketIndex));
                    }
                }
            }

            // Handle simple property
            int nextDot = propertyPath.IndexOf('.');
            if (nextDot == -1)
            {
                // Last property in chain
                return GetPropertyValue(obj, propertyPath);
            }
            else
            {
                // More properties to traverse
                string currentProp = propertyPath.Substring(0, nextDot);
                string remainingPath = propertyPath.Substring(nextDot + 1);

                obj = GetPropertyValue(obj, currentProp);
                return obj != null ? ResolveNestedProperty(obj, remainingPath) : null;
            }
        }

        /// <summary>
        /// Gets a property value using reflection
        /// </summary>
        private object GetPropertyValue(object obj, string propertyName)
        {
            if (obj == null)
                return null;

            var type = obj.GetType();

            // Check for dictionary-like access
            if (obj is IDictionary dictionary && dictionary.Contains(propertyName))
                return dictionary[propertyName];

            // Try to get a property with the given name
            var property = type.GetProperty(propertyName);
            if (property != null)
                return property.GetValue(obj, null);

            // Try to get a field with the given name
            var field = type.GetField(propertyName);
            if (field != null)
                return field.GetValue(obj);

            // If we get here, the property/field wasn't found
            return null;
        }

        /// <summary>
        /// Resolves an array indexer in the format varName[index]
        /// </summary>
        private object ResolveArrayIndexer(string expression)
        {
            if (string.IsNullOrEmpty(expression))
                return null;

            int openBracket = expression.IndexOf('[');
            if (openBracket == -1)
                return null;

            int closeBracket = expression.IndexOf(']', openBracket);
            if (closeBracket == -1)
                return null;

            string variableName = expression.Substring(0, openBracket);
            string indexStr = expression.Substring(openBracket + 1, closeBracket - openBracket - 1);

            if (!_variables.TryGetValue(variableName, out var value) || value == null)
                return null;

            if (!(value is IList list))
                return null;

            if (!int.TryParse(indexStr, out int index) || index < 0 || index >= list.Count)
                return null;

            var itemValue = list[index];

            // Check if there's more to the path
            if (closeBracket + 1 < expression.Length)
            {
                if (expression[closeBracket + 1] == '.')
                {
                    // Continue with nested property
                    string propertyPath = expression.Substring(closeBracket + 2);
                    return ResolveNestedProperty(itemValue, propertyPath);
                }
            }

            return itemValue;
        }

        /// <summary>
        /// Tries to get a collection and its count
        /// </summary>
        internal bool TryGetCollection(string collectionName, out ICollection collection)
        {
            collection = null;

            // Check variables
            if (_variables.TryGetValue(collectionName, out var value))
            {
                // Direct ICollection
                if (value is ICollection col)
                {
                    collection = col;
                    return true;
                }

                if (value is IEnumerable enumerable)
                {
                    var list = new List<object>();

                    // Create a safe snapshot of the enumerable to avoid collection modified exceptions
                    var snapshot = enumerable.Cast<object>().ToArray();
                    foreach (var item in snapshot)
                    {
                        list.Add(item);
                    }

                    collection = list;
                    return true;
                }
            }

            return false;
        }
    }
}