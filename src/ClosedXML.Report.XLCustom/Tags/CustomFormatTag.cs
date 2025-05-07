using ClosedXML.Report.Options;

namespace ClosedXML.Report.XLCustom.Tags;

public class CustomFormatTag : OptionTag
{
    /// <summary>
    /// Executes the tag processing
    /// </summary>
    public override void Execute(ProcessingContext context)
    {
        var xlCell = Cell.GetXlCell(context.Range);
        var formatRegistry = XLCustomRegistry.Instance.FormatRegistry;

        try
        {
            // 이름 기반 매개변수 가져오기
            var variableName = GetParameter("name");
            var formatName = GetParameter("format");

            Log.Debug($"CustomFormatTag - variable: {variableName}, format: {formatName}");

            if (string.IsNullOrEmpty(variableName) || string.IsNullOrEmpty(formatName))
            {
                xlCell.Value = "Invalid format expression";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // Evaluate variable
            var value = context.Evaluator.Evaluate(variableName, new Parameter("item", context.Value));
            Log.Debug($"Evaluated variable {variableName} = {value ?? "null"}");

            // Check format registry first
            if (formatRegistry.IsRegistered(formatName))
            {
                // Apply format
                var formatted = formatRegistry.ApplyFormat(formatName, value, Array.Empty<string>());
                xlCell.SetValue(formatted);
                Log.Debug($"Applied format {formatName} = {formatted ?? "null"}");
            }
            else
            {
                xlCell.Value = $"Unknown format: {formatName}";
                xlCell.Style.Font.FontColor = XLColor.Red;
                Log.Debug($"Format not found: {formatName}");
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"Error in CustomFormatTag: {ex.Message}");
            xlCell.Value = $"Error: {ex.Message}";
            xlCell.Style.Font.FontColor = XLColor.Red;
        }
    }
}