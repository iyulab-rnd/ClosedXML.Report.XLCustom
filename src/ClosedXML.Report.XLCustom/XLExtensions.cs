namespace ClosedXML.Report.XLCustom;

public static class XLExtensions
{
    public static void SetValue(this IXLCell cell, object value)
    {
        var cellValue = XLCellValueConverter.FromObject(value);
        cell.SetValue(cellValue);
    }
}
