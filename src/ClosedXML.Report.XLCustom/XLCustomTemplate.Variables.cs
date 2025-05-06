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

        // Store variable in our internal dictionary
        _variables[alias] = value;

        // Handle DataTable conversion for collections
        if (value is DataTable dt)
            value = dt.Rows.Cast<DataRow>();

        // Add to all interpreters
        foreach (var interpreter in _interpreters.Values)
        {
            interpreter.AddVariable(alias, value);
        }

        // Only add to base template if not already added
        if (!_variablesAddedToBase.Contains(alias))
        {
            _baseTemplate.AddVariable(alias, value);
            _variablesAddedToBase.Add(alias);
        }
    }

    /// <summary>
    /// Resolves a variable directly
    /// </summary>
    internal object ResolveVariable(string variableName)
    {
        // 내부 변수 딕셔너리에서 찾기
        if (_variables.TryGetValue(variableName, out var value))
            return value;

        // 글로벌 리졸버 시도
        if (TryResolveGlobal(variableName, out value))
            return value;

        return null;
    }
}