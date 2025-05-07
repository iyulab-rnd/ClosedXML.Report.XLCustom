namespace ClosedXML.Report.XLCustom;

public static class XLExtensions
{
    public static void SetValue(this IXLCell cell, object value)
    {
        cell.Value = XLCellValueConverter.FromObject(value);
    }
}
