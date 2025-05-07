namespace ClosedXML.Report.XLCustom;

/// <summary>
/// 레지스트리에 접근하기 위한 싱글톤 클래스
/// </summary>
public class XLCustomRegistry
{
    private static XLCustomRegistry _instance;
    private static readonly object _lock = new object();

    // 테스트를 위한 리셋 메서드 추가
    public static void Reset()
    {
        lock (_lock)
        {
            _instance = null;
        }
    }

    public static XLCustomRegistry Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new XLCustomRegistry();
                    }
                }
            }
            return _instance;
        }
    }

    public FormatRegistry FormatRegistry { get; private set; }
    public FunctionRegistry FunctionRegistry { get; private set; }

    private XLCustomRegistry()
    {
        FormatRegistry = new FormatRegistry();
        FunctionRegistry = new FunctionRegistry();
    }

    public void SetRegistries(FormatRegistry formatRegistry, FunctionRegistry functionRegistry)
    {
        Log.Debug("Setting new registries in XLCustomRegistry singleton");
        FormatRegistry = formatRegistry ?? throw new ArgumentNullException(nameof(formatRegistry));
        FunctionRegistry = functionRegistry ?? throw new ArgumentNullException(nameof(functionRegistry));
    }
}