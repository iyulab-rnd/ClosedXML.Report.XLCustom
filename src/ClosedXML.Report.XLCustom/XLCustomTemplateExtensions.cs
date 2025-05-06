namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Extension methods for XLCustomTemplate
/// </summary>
public static class XLCustomTemplateExtensions
{
    /// <summary>
    /// Registers all built-in formatters with the template
    /// </summary>
    public static XLCustomTemplate RegisterBuiltInFormatters(this XLCustomTemplate template)
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));

        Debug.WriteLine("Registering built-in formatters");
        template.RegisterFormat("upper", BuiltInFormatters.Upper);
        template.RegisterFormat("lower", BuiltInFormatters.Lower);
        template.RegisterFormat("titlecase", BuiltInFormatters.TitleCase);
        template.RegisterFormat("mask", BuiltInFormatters.Mask);
        template.RegisterFormat("phone", BuiltInFormatters.Phone);
        template.RegisterFormat("truncate", BuiltInFormatters.Truncate);

        return template;
    }

    /// <summary>
    /// Registers all built-in functions with the template
    /// </summary>
    public static XLCustomTemplate RegisterBuiltInFunctions(this XLCustomTemplate template)
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));

        Debug.WriteLine("Registering built-in functions");
        // 함수 이름을 소문자로 변경
        template.RegisterFunction("bold", BuiltInFunctions.Bold);
        template.RegisterFunction("italic", BuiltInFunctions.Italic);
        template.RegisterFunction("color", BuiltInFunctions.Color);
        template.RegisterFunction("link", BuiltInFunctions.Link);
        template.RegisterFunction("image", BuiltInFunctions.Image);

        Debug.WriteLine("Built-in functions registered");
        return template;
    }

    /// <summary>
    /// Registers all built-in formatters and functions with the template
    /// </summary>
    public static XLCustomTemplate RegisterBuiltIns(this XLCustomTemplate template)
    {
        Debug.WriteLine("Registering all built-ins");
        return template.RegisterBuiltInFormatters().RegisterBuiltInFunctions();
    }
}