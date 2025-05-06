namespace ClosedXML.Report.XLCustom;

public partial class XLCustomTemplate
{
    /// <summary>
    /// Evaluates a formula expression with variables
    /// </summary>
    public object EvaluateFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return null;

        Debug.WriteLine($"Evaluating formula: {formula}");

        try
        {
            // 간단한 사칙연산 처리
            if (formula.Contains(" * "))
            {
                return EvaluateMultiplication(formula);
            }
            else if (formula.Contains(" / "))
            {
                return EvaluateDivision(formula);
            }
            else if (formula.Contains(" + "))
            {
                return EvaluateAddition(formula);
            }
            else if (formula.Contains(" - "))
            {
                return EvaluateSubtraction(formula);
            }

            // 단일 변수 참조
            return ResolveVariable(formula);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Formula evaluation error: {ex.Message}");
            return formula; // 에러 시 원본 반환
        }
    }

    /// <summary>
    /// Evaluates multiplication expressions
    /// </summary>
    private object EvaluateMultiplication(string formula)
    {
        string[] parts = formula.Split(new[] { " * " }, StringSplitOptions.None);
        if (parts.Length != 2)
            return formula;

        string leftPart = parts[0].Trim();
        string rightPart = parts[1].Trim();

        // 좌변 값 평가
        object leftValue = null;
        if (leftPart.Contains("."))
        {
            leftValue = ResolveVariable(leftPart);
        }
        else if (decimal.TryParse(leftPart, out decimal leftConstant))
        {
            leftValue = leftConstant;
        }

        // 우변 값 평가
        object rightValue = null;
        if (rightPart.Contains("."))
        {
            rightValue = ResolveVariable(rightPart);
        }
        else if (decimal.TryParse(rightPart, out decimal rightConstant))
        {
            rightValue = rightConstant;
        }

        // 계산 수행
        if (leftValue != null && rightValue != null)
        {
            if (decimal.TryParse(leftValue.ToString(), out decimal leftDecimal) &&
                decimal.TryParse(rightValue.ToString(), out decimal rightDecimal))
            {
                return leftDecimal * rightDecimal;
            }
        }

        return formula; // 처리할 수 없는 경우 원본 반환
    }

    /// <summary>
    /// Evaluates division expressions
    /// </summary>
    private object EvaluateDivision(string formula)
    {
        string[] parts = formula.Split(new[] { " / " }, StringSplitOptions.None);
        if (parts.Length != 2)
            return formula;

        string leftPart = parts[0].Trim();
        string rightPart = parts[1].Trim();

        // 좌변 값 평가
        object leftValue = null;
        if (leftPart.Contains("."))
        {
            leftValue = ResolveVariable(leftPart);
        }
        else if (decimal.TryParse(leftPart, out decimal leftConstant))
        {
            leftValue = leftConstant;
        }

        // 우변 값 평가
        object rightValue = null;
        if (rightPart.Contains("."))
        {
            rightValue = ResolveVariable(rightPart);
        }
        else if (decimal.TryParse(rightPart, out decimal rightConstant))
        {
            rightValue = rightConstant;
        }

        // 계산 수행
        if (leftValue != null && rightValue != null)
        {
            if (decimal.TryParse(leftValue.ToString(), out decimal leftDecimal) &&
                decimal.TryParse(rightValue.ToString(), out decimal rightDecimal) &&
                rightDecimal != 0)
            {
                return leftDecimal / rightDecimal;
            }
        }

        return formula; // 처리할 수 없는 경우 원본 반환
    }

    /// <summary>
    /// Evaluates addition expressions
    /// </summary>
    private object EvaluateAddition(string formula)
    {
        string[] parts = formula.Split(new[] { " + " }, StringSplitOptions.None);
        if (parts.Length != 2)
            return formula;

        string leftPart = parts[0].Trim();
        string rightPart = parts[1].Trim();

        // 좌변 값 평가
        object leftValue = null;
        if (leftPart.Contains("."))
        {
            leftValue = ResolveVariable(leftPart);
        }
        else if (decimal.TryParse(leftPart, out decimal leftConstant))
        {
            leftValue = leftConstant;
        }

        // 우변 값 평가
        object rightValue = null;
        if (rightPart.Contains("."))
        {
            rightValue = ResolveVariable(rightPart);
        }
        else if (decimal.TryParse(rightPart, out decimal rightConstant))
        {
            rightValue = rightConstant;
        }

        // 계산 수행
        if (leftValue != null && rightValue != null)
        {
            // 문자열 연결
            if (leftValue is string || rightValue is string)
            {
                return leftValue?.ToString() + rightValue?.ToString();
            }

            // 숫자 덧셈
            if (decimal.TryParse(leftValue.ToString(), out decimal leftDecimal) &&
                decimal.TryParse(rightValue.ToString(), out decimal rightDecimal))
            {
                return leftDecimal + rightDecimal;
            }
        }

        return formula; // 처리할 수 없는 경우 원본 반환
    }

    /// <summary>
    /// Evaluates subtraction expressions
    /// </summary>
    private object EvaluateSubtraction(string formula)
    {
        string[] parts = formula.Split(new[] { " - " }, StringSplitOptions.None);
        if (parts.Length != 2)
            return formula;

        string leftPart = parts[0].Trim();
        string rightPart = parts[1].Trim();

        // 좌변 값 평가
        object leftValue = null;
        if (leftPart.Contains("."))
        {
            leftValue = ResolveVariable(leftPart);
        }
        else if (decimal.TryParse(leftPart, out decimal leftConstant))
        {
            leftValue = leftConstant;
        }

        // 우변 값 평가
        object rightValue = null;
        if (rightPart.Contains("."))
        {
            rightValue = ResolveVariable(rightPart);
        }
        else if (decimal.TryParse(rightPart, out decimal rightConstant))
        {
            rightValue = rightConstant;
        }

        // 계산 수행
        if (leftValue != null && rightValue != null)
        {
            if (decimal.TryParse(leftValue.ToString(), out decimal leftDecimal) &&
                decimal.TryParse(rightValue.ToString(), out decimal rightDecimal))
            {
                return leftDecimal - rightDecimal;
            }
        }

        return formula; // 처리할 수 없는 경우 원본 반환
    }
}