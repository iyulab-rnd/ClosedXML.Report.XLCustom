namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Partial class containing function registration functionality
/// </summary>
public partial class XLCustomTemplate
{
    /// <summary>
    /// Registers a custom function processor that can manipulate cells
    /// </summary>
    public XLCustomTemplate RegisterFunction(string functionName, XLFunctionHandler function)
    {
        _functionRegistry.Register(functionName, function);
        _preprocessed = false; // Need to reprocess after registering functions

        // 등록 확인 로깅
        Log.Debug($"Function '{functionName}' registered in {(_useGlobalRegistry ? "global" : "local")} registry");
        return this;
    }

    /// <summary>
    /// Registers all built-in functions and global variables
    /// </summary>
    public XLCustomTemplate RegisterBuiltIns()
    {
        RegisterBuiltInFunctions();
        RegisterBuiltInGlobalVariables();
        return this;
    }

    /// <summary>
    /// Registers built-in functions
    /// </summary>
    public XLCustomTemplate RegisterBuiltInFunctions()
    {
        // 내장 함수 등록을 FunctionRegistry 클래스로 이동
        if (_useGlobalRegistry)
        {
            // 전역 레지스트리 사용 시 해당 레지스트리에 등록
            XLCustomRegistry.Instance.FunctionRegistry.RegisterBuiltInFunctions();
        }
        else
        {
            // 로컬 레지스트리 사용 시 해당 레지스트리에 등록
            (_functionRegistry as FunctionRegistry)?.RegisterBuiltInFunctions();
        }

        // 전처리 플래그 재설정
        _preprocessed = false;

        // 등록된 함수 로깅
        Log.Debug("Built-in functions registered");
        foreach (var name in _functionRegistry.GetFunctionNames())
        {
            Log.Debug($"Registered function: {name}");
        }

        return this;
    }

    /// <summary>
    /// Debug method: Process an expression and return the result
    /// </summary>
    public string DebugExpression(string expression, IXLCell cell = null)
    {
        var expressionProcessor = new XLExpressionProcessor(_functionRegistry, _globalVariables);
        return expressionProcessor.ProcessExpression(expression, cell);
    }

    /// <summary>
    /// Gets the function registry for testing purposes
    /// </summary>
    public FunctionRegistry GetFunctionRegistryForTest()
    {
        return _functionRegistry as FunctionRegistry;
    }
}