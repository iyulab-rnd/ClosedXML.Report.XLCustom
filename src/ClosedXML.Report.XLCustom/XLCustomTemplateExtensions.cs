namespace ClosedXML.Report.XLCustom
{
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
            template.RegisterFormat("upper", BuiltInFormatters.Upper);
            template.RegisterFormat("lower", BuiltInFormatters.Lower);
            template.RegisterFormat("titlecase", BuiltInFormatters.TitleCase);
            template.RegisterFormat("mask", BuiltInFormatters.Mask);
            template.RegisterFormat("truncate", BuiltInFormatters.Truncate);
            template.RegisterFormat("currency", BuiltInFormatters.Currency);
            template.RegisterFormat("number", BuiltInFormatters.Number);
            template.RegisterFormat("percent", BuiltInFormatters.Percent);
            template.RegisterFormat("date", BuiltInFormatters.Date);

            return template;
        }

        /// <summary>
        /// Registers all built-in functions with the template
        /// </summary>
        public static XLCustomTemplate RegisterBuiltInFunctions(this XLCustomTemplate template)
        {
            template.RegisterFunction("bold", BuiltInFunctions.Bold);
            template.RegisterFunction("italic", BuiltInFunctions.Italic);
            template.RegisterFunction("color", BuiltInFunctions.Color);
            template.RegisterFunction("link", BuiltInFunctions.Link);
            template.RegisterFunction("image", BuiltInFunctions.Image);
            template.RegisterFunction("format", BuiltInFunctions.Format);
            template.RegisterFunction("background", BuiltInFunctions.BackgroundColor);
            template.RegisterFunction("center", BuiltInFunctions.Center);
            template.RegisterFunction("border", BuiltInFunctions.Border);

            return template;
        }

        /// <summary>
        /// Registers all built-in formatters and functions
        /// </summary>
        public static XLCustomTemplate RegisterBuiltIns(this XLCustomTemplate template)
        {
            return template.RegisterBuiltInFormatters().RegisterBuiltInFunctions();
        }

        /// <summary>
        /// Adds a custom variable resolver that will be used when a variable is not found
        /// </summary>
        public static XLCustomTemplate WithVariableResolver(this XLCustomTemplate template, Func<string, object> resolver)
        {
            template.SetGlobalResolver(resolver);
            return template;
        }

        /// <summary>
        /// Adds a model object to the template with a specific alias
        /// </summary>
        public static XLCustomTemplate WithModel(this XLCustomTemplate template, string alias, object model)
        {
            template.AddVariable(alias, model);
            return template;
        }

        /// <summary>
        /// Adds a model object to the template
        /// </summary>
        public static XLCustomTemplate WithModel(this XLCustomTemplate template, object model)
        {
            template.AddVariable(model);
            return template;
        }

        /// <summary>
        /// Updates a variable with a new value, overwriting any existing variable with the same name
        /// </summary>
        public static XLCustomTemplate UpdateVariable(this XLCustomTemplate template, string alias, object value)
        {
            template.AddVariable(alias, value); // AddVariable will handle overwriting
            return template;
        }

        /// <summary>
        /// Updates multiple variables from a dictionary
        /// </summary>
        public static XLCustomTemplate UpdateVariables(this XLCustomTemplate template, IDictionary<string, object> variables)
        {
            foreach (var kvp in variables)
            {
                template.AddVariable(kvp.Key, kvp.Value);
            }
            return template;
        }
    }
}