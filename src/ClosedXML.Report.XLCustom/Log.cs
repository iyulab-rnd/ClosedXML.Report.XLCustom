using System.Diagnostics;

namespace ClosedXML.Report.XLCustom;

internal class Log
{
    [Conditional("DEBUG")]
    public static void Debug(string message)
    {
        System.Diagnostics.Debug.WriteLine(message);
    }
}
