using ClosedXML.Excel;
using ClosedXML.Report.Options;
using System;
using System.Globalization;

namespace ClosedXML.Report.XLCustom.Tags;

public class FormatTag : OptionTag
{
    public override void Execute(ProcessingContext context)
    {
        var xlCell = Cell.GetXlCell(context.Range);

        try
        {
            var variableName = GetParameter("name");
            var formatString = GetParameter("format");

            Log.Debug($"FormatTag - name: {variableName}, format: {formatString}");

            if (string.IsNullOrEmpty(variableName) || string.IsNullOrEmpty(formatString))
            {
                Log.Debug("FormatTag - Missing required parameters");
                return;
            }

            // 변수 평가
            var value = context.Evaluator.Evaluate(variableName, new Parameter("item", context.Value));
            Log.Debug($"Evaluated variable {variableName} = {value ?? "null"}");

            // 값이 null이면 처리하지 않음
            if (value == null)
            {
                xlCell.SetValue(Blank.Value);
                return;
            }

            // 포맷팅된 값을 직접 생성
            string formattedValue = FormatValue(value, formatString);

            // 포맷팅된 값을 셀에 직접 할당
            xlCell.Value = formattedValue;

            Log.Debug($"Set formatted value directly: {formattedValue}");
        }
        catch (Exception ex)
        {
            Log.Debug($"Error in FormatTag: {ex.Message}");
            xlCell.Value = $"Error: {ex.Message}";
            xlCell.Style.Font.FontColor = XLColor.Red;
        }
    }

    /// <summary>
    /// Formats a value using the specified format string
    /// </summary>
    private string FormatValue(object value, string formatString)
    {
        // DateTime 값 처리
        if (value is DateTime dateTime)
        {
            return dateTime.ToString(formatString, CultureInfo.CurrentCulture);
        }

        // 숫자 값 처리
        if (value is IFormattable formattable)
        {
            return formattable.ToString(formatString, CultureInfo.CurrentCulture);
        }

        // 다른 타입의 값은 ToString() 호출
        return value.ToString();
    }
}