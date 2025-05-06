namespace ClosedXML.Report.XLCustom;

public partial class XLCustomTemplate
{
    /// <summary>
    /// Adds variables from an object
    /// </summary>
    public void AddVariable(object value)
    {
        CheckIsDisposed();

        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                AddVariable(entry.Key.ToString(), entry.Value);
            }
        }
        else
        {
            var type = value.GetType();
            var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance)
                .Where(f => f.IsPublic)
                .Select(f => new { f.Name, val = f.GetValue(value), type = f.FieldType })
                .Concat(
                    type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                        .Where(f => f.CanRead)
                        .Select(f => new {
                            f.Name,
                            val = f.GetValue(value, Array.Empty<object>()),
                            type = f.PropertyType
                        })
                );

            foreach (var field in fields)
            {
                AddVariable(field.Name, field.val);
            }
        }
    }

    /// <summary>
    /// Adds a variable with a specific name
    /// </summary>
    public void AddVariable(string alias, object value)
    {
        CheckIsDisposed();

        Debug.WriteLine($"Adding variable: {alias} = {value}");

        // 중복 체크
        if (_variables.ContainsKey(alias))
        {
            Debug.WriteLine($"Variable {alias} already exists, updating value");
            _variables[alias] = value;
        }
        else
        {
            // Store variable in our internal dictionary
            _variables.Add(alias, value);
        }

        // Handle DataTable conversion for collections
        object baseValue = value;
        if (value is DataTable dt)
            baseValue = dt.Rows.Cast<DataRow>();

        // Add to all interpreters
        foreach (var interpreter in _interpreters.Values)
        {
            interpreter.AddVariable(alias, baseValue);
        }

        // 기본 템플릿에 추가
        try
        {
            _baseTemplate.AddVariable(alias, baseValue);

            // Record that it was added
            if (!_variablesAddedToBase.Contains(alias))
            {
                _variablesAddedToBase.Add(alias);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error adding variable to base template: {ex.Message}");
        }
    }

    /// <summary>
    /// Resolves a variable directly
    /// </summary>
    internal object ResolveVariable(string variableName)
    {
        Debug.WriteLine($"Resolving variable: {variableName}");

        // 내부 변수 딕셔너리에서 찾기
        if (_variables.TryGetValue(variableName, out var value))
        {
            Debug.WriteLine($"Found in variables: {value}");
            return value;
        }

        // 글로벌 리졸버 시도
        if (TryResolveGlobal(variableName, out value))
        {
            Debug.WriteLine($"Found in global resolver: {value}");
            return value;
        }

        // 컬렉션 확인
        if (variableName.Contains("_Count"))
        {
            string collectionName = variableName.Substring(0, variableName.Length - 6);
            if (TryGetCollection(collectionName, out var collection))
            {
                Debug.WriteLine($"Found collection count: {collection.Count}");
                return collection.Count;
            }
        }

        Debug.WriteLine($"Variable not found: {variableName}");
        return null;
    }

    /// <summary>
    /// Tries to get a collection and its count
    /// </summary>
    internal bool TryGetCollection(string collectionName, out ICollection collection)
    {
        collection = null;

        Debug.WriteLine($"Looking for collection: {collectionName}");

        // Check variables
        if (_variables.TryGetValue(collectionName, out var value))
        {
            // Direct ICollection
            if (value is ICollection col)
            {
                Debug.WriteLine($"Found as ICollection with {col.Count} items");
                collection = col;
                return true;
            }

            // Convert IEnumerable to List
            if (value is IEnumerable enumerable)
            {
                var list = new List<object>();
                foreach (var item in enumerable)
                {
                    list.Add(item);
                }

                Debug.WriteLine($"Converted IEnumerable to List with {list.Count} items");
                collection = list;
                return true;
            }
        }

        Debug.WriteLine($"Collection not found: {collectionName}");
        return false;
    }
}