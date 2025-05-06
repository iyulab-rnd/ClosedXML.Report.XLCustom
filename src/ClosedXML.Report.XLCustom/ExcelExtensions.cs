namespace ClosedXML.Report.XLCustom;

public static class ExcelExtensions
{
    public static void SetValue(this IXLCell cell, object value)
    {
        var cellValue = XLCellValueConverter.FromObject(value);
        cell.SetValue(cellValue);
    }
}
