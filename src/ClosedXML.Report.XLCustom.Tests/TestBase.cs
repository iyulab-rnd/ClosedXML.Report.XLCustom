using System.Diagnostics;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests;

/// <summary>
/// Base class for all test classes providing common functionality
/// </summary>
public abstract class TestBase
{
    protected readonly ITestOutputHelper Output;
    private readonly XUnitTraceListener _listener;

    public TestBase(ITestOutputHelper output)
    {
        Output = output;
        _listener = new XUnitTraceListener(output);

        // 모든 테스트에서 공유되지 않도록 리스너 추가
        Trace.Listeners.Add(_listener);
    }

    // 테스트 클래스에서 리소스 정리를 위한 Dispose 패턴 구현
    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            // 리스너 제거하여 다른 테스트에 영향 없도록
            Trace.Listeners.Remove(_listener);
        }
    }

    protected void LogResult(XLGenerateResult result)
    {
        if (result.HasErrors)
        {
            foreach (var error in result.ParsingErrors)
            {
                Output.WriteLine($"Error: {error.Message}, Range: {error.Range}");
            }
        }
    }
}

/// <summary>
/// Trace listener that redirects output to XUnit's test output
/// </summary>
public class XUnitTraceListener : TraceListener
{
    private readonly ITestOutputHelper _output;

    public XUnitTraceListener(ITestOutputHelper output)
    {
        _output = output;
    }

    public override void Write(string message)
    {
        try
        {
            _output.WriteLine(message);
        }
        catch (Exception)
        {
            // 테스트 컨텍스트 외부에서 호출되는 경우 무시
            // 필요하다면 여기서 콘솔이나 다른 로그 메커니즘으로 출력 가능
            Console.Write(message);
        }
    }

    public override void WriteLine(string message)
    {
        try
        {
            _output.WriteLine(message);
        }
        catch (Exception)
        {
            // 테스트 컨텍스트 외부에서 호출되는 경우 무시
            // 필요하다면 여기서 콘솔이나 다른 로그 메커니즘으로 출력 가능
            Console.WriteLine(message);
        }
    }
}