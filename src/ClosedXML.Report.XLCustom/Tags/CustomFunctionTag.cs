using ClosedXML.Report.Options;

namespace ClosedXML.Report.XLCustom.Tags;

public class CustomFunctionTag : OptionTag
{
    public override void Execute(ProcessingContext context)
    {
        var xlCell = Cell.GetXlCell(context.Range);
        var functionRegistry = XLCustomRegistry.Instance.FunctionRegistry;

        try
        {
            // 이름 기반 매개변수 가져오기
            var variableName = GetParameter("name");
            var functionName = GetParameter("function");
            var parametersStr = GetParameter("parameters");

            // 파라미터 문자열 파싱
            var parameters = new List<string>();
            if (!string.IsNullOrEmpty(parametersStr))
            {
                parameters.AddRange(parametersStr.Split(',').Select(p => UnescapeParameter(p)));
            }

            Log.Debug($"CustomFunctionTag - variable: {variableName}, function: {functionName}, params: {string.Join(", ", parameters)}");

            if (string.IsNullOrEmpty(variableName) || string.IsNullOrEmpty(functionName))
            {
                xlCell.Value = "Invalid function expression";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // Evaluate variable
            var value = context.Evaluator.Evaluate(variableName, new Parameter("item", context.Value));
            Log.Debug($"Evaluated variable {variableName} = {value ?? "null"}");

            // Apply function
            if (functionRegistry.IsRegistered(functionName))
            {
                var function = functionRegistry.GetFunction(functionName);
                function(xlCell, value, parameters.ToArray());
                Log.Debug($"Applied function {functionName}");
            }
            else
            {
                xlCell.Value = $"Unknown function: {functionName}";
                xlCell.Style.Font.FontColor = XLColor.Red;
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"Error in CustomFunctionTag: {ex.Message}");
            xlCell.Value = $"Error: {ex.Message}";
            xlCell.Style.Font.FontColor = XLColor.Red;
        }
    }

    /// <summary>
    /// Unescapes a parameter value from tag parameters
    /// </summary>
    private string UnescapeParameter(string param)
    {
        // 홑따옴표로 묶인 매개변수 처리
        if (param.StartsWith("'") && param.EndsWith("'") && param.Length >= 2)
        {
            // 따옴표 제거 및 이스케이프된 따옴표 처리
            param = param.Substring(1, param.Length - 2).Replace("''", "'");
        }

        return param;
    }
}