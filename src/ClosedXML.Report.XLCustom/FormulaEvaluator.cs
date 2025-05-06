using System.Linq.Expressions;
using System.Text.RegularExpressions;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Evaluates formulas in templates
/// </summary>
public class FormulaEvaluator
{
    private static readonly Regex ExprMatch = new Regex(@"\{\{.+?\}\}");

    private readonly Dictionary<string, Delegate> _lambdaCache = new Dictionary<string, Delegate>();
    private readonly Dictionary<string, object> _variables = new Dictionary<string, object>();
    private Func<string, object> _globalResolver;

    /// <summary>
    /// Evaluates a formula
    /// </summary>
    public object Evaluate(string formula, params Parameter[] pars)
    {
        if (string.IsNullOrEmpty(formula))
            return formula;

        var expressions = GetExpressions(formula);
        foreach (var expr in expressions)
        {
            var val = Eval(Trim(expr), pars);
            if (expr == formula)
                return val;

            formula = formula.Replace(expr, val?.ToString() ?? string.Empty);
        }
        return formula;
    }

    /// <summary>
    /// Tries to evaluate a formula
    /// </summary>
    public bool TryEvaluate(string formula, out object result, params Parameter[] pars)
    {
        try
        {
            result = Evaluate(formula, pars);
            return true;
        }
        catch
        {
            result = null;
            return false;
        }
    }

    /// <summary>
    /// Adds a variable
    /// </summary>
    public void AddVariable(string name, object value)
    {
        _variables[name] = value;
    }

    /// <summary>
    /// Sets a global resolver function
    /// </summary>
    public void SetGlobalResolver(Func<string, object> resolver)
    {
        _globalResolver = resolver;
    }

    /// <summary>
    /// Gets the expressions from a formula
    /// </summary>
    private IEnumerable<string> GetExpressions(string cellValue)
    {
        var matches = ExprMatch.Matches(cellValue);
        if (matches.Count == 0)
            return new[] { cellValue };
        return from Match match in matches select match.Value;
    }

    /// <summary>
    /// Trims the delimiters from an expression
    /// </summary>
    private string Trim(string formula)
    {
        if (formula.StartsWith("{{"))
            return formula.Substring(2, formula.Length - 4);
        else
            return formula;
    }

    /// <summary>
    /// Tries to resolve a variable from the global resolver
    /// </summary>
    public bool TryResolveGlobal(string variableName, out object value)
    {
        if (_globalResolver != null)
        {
            try
            {
                value = _globalResolver(variableName);
                return value != null;
            }
            catch
            {
                value = null;
                return false;
            }
        }

        value = null;
        return false;
    }

    /// <summary>
    /// Evaluates a formula
    /// </summary>
    private object Eval(string expression, Parameter[] pars)
    {
        // Check variables first
        if (_variables.TryGetValue(expression, out var variable))
            return variable;

        // Then check parameters
        foreach (var par in pars)
        {
            if (par.ParameterExpression.Name == expression)
                return par.Value;
        }

        // Try global resolver
        if (TryResolveGlobal(expression, out var resolved))
            return resolved;

        // Could not resolve
        return expression;
    }

    /// <summary>
    /// Parses an expression into a lambda
    /// </summary>
    internal Delegate ParseExpression(string formula, ParameterExpression[] parameters)
    {
        var cacheKey = GetCacheKey(formula, parameters);
        if (!_lambdaCache.TryGetValue(cacheKey, out var lambda))
        {
            try
            {
                // Here you would use a parser library like DynamicLinq or write your own parser
                // This is a simplified placeholder for the actual parsing logic
                throw new NotImplementedException("Expression parsing not implemented");
            }
            catch (ArgumentException)
            {
                return null;
            }

            _lambdaCache.Add(cacheKey, lambda);
        }
        return lambda;
    }

    private string GetCacheKey(string formula, ParameterExpression[] parameters)
    {
        return formula + string.Join("+", parameters.Select(x => x.Type.Name));
    }
}

/// <summary>
/// Parameter for formula evaluation
/// </summary>
public class Parameter
{
    /// <summary>
    /// Gets the parameter expression
    /// </summary>
    public ParameterExpression ParameterExpression { get; }

    /// <summary>
    /// Gets the parameter value
    /// </summary>
    public object Value { get; }

    /// <summary>
    /// Creates a new parameter
    /// </summary>
    public Parameter(string name, object value)
    {
        ParameterExpression = Expression.Parameter(value?.GetType() ?? typeof(object), name);
        Value = value;
    }
}