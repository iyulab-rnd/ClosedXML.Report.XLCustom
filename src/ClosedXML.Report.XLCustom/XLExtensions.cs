using ClosedXML.Report.Utils;

namespace ClosedXML.Report.XLCustom
{
    public static class XLExtensions
    {
        public static void SetValue(this IXLCell cell, object value)
        {
            // DateTime을 포함한 기본 타입 처리
            if (value is DateTime dateValue)
            {
                cell.Value = dateValue;
            }
            else if (value is TimeSpan timeValue)
            {
                cell.Value = timeValue;
            }
            else
            {
                cell.SetValue(XLCellValueConverter.FromObject(value));
            }
        }
    }
}