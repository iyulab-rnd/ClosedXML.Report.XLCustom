namespace ClosedXML.Report.XLCustom.Internals;

/// <summary>
/// Custom range interpreter that extends the standard RangeInterpreter
/// </summary>
internal class CustomRangeInterpreter : RangeInterpreter
{
    private readonly XLCustomTemplate _template;
    private readonly TemplateErrors _errors;
    private FormulaEvaluator _baseEvaluator;

    /// <summary>
    /// Gets the base formula evaluator
    /// </summary>
    public FormulaEvaluator BaseEvaluator
    {
        get
        {
            if (_baseEvaluator == null)
            {
                // Get base evaluator via reflection
                var fieldInfo = typeof(RangeInterpreter).GetField("_evaluator",
                    BindingFlags.NonPublic | BindingFlags.Instance);

                if (fieldInfo != null)
                {
                    _baseEvaluator = (FormulaEvaluator)fieldInfo.GetValue(this);
                }
                else
                {
                    throw new InvalidOperationException(
                        "Could not access base evaluator. Library may be incompatible with this version of ClosedXML.Report.");
                }
            }

            return _baseEvaluator;
        }
    }

    public CustomRangeInterpreter(string alias, TemplateErrors errors, XLCustomTemplate template)
        : base(alias, errors)
    {
        _template = template ?? throw new ArgumentNullException(nameof(template));
        _errors = errors ?? throw new ArgumentNullException(nameof(errors));
    }

    // Override value evaluation - core functionality
    public override void EvaluateValues(IXLRange range, params Parameter[] pars)
    {
        try
        {
            // First check if variables exist in the global resolver
            foreach (var par in pars)
            {
                if (_template.TryResolveGlobal(par.ParameterExpression.Name, out var value))
                {
                    AddVariable(par.ParameterExpression.Name, value);
                }
            }

            // Call base implementation - handles standard ClosedXML.Report expressions
            base.EvaluateValues(range, pars);

            // Process enhanced expressions
            var customEvaluator = new CustomFormulaEvaluator(_template, BaseEvaluator);

            // Find cells with enhanced expressions
            var enhancedCells = range.CellsUsed(cell =>
            {
                var value = cell.GetString();
                return value.Contains("{{") && (value.Contains(":") || value.Contains("|"));
            });

            foreach (var cell in enhancedCells)
            {
                string cellValue = cell.GetString();

                // Check if cell still contains enhanced expressions
                // (after base processing)
                if (cellValue.Contains("{{") && (cellValue.Contains(":") || cellValue.Contains("|")))
                {
                    try
                    {
                        // Evaluate expression
                        var result = customEvaluator.Evaluate(cellValue, cell, pars);

                        // Set value
                        if (result != null)
                        {
                            if (cellValue.StartsWith("&="))
                            {
                                // Formula should be set as FormulaA1
                                cell.FormulaA1 = result.ToString();
                            }
                            else
                            {
                                // Regular value
                                cell.SetValue(result);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Show error
                        _errors.Add(new TemplateError(
                            $"Error processing enhanced expression in cell {cell.Address}: {ex.Message}",
                            range.Worksheet.Range(cell.Address, cell.Address)));

                        cell.Value = $"Error: {ex.Message}";
                        cell.Style.Font.FontColor = XLColor.Red;
                    }
                }
            }

            // Process collection metadata expressions - like Count
            ProcessCollectionMetadata(range, pars);
        }
        catch (Exception ex)
        {
            _errors.Add(new TemplateError($"Error evaluating range: {ex.Message}", range));
        }
    }

    /// <summary>
    /// Process collection metadata like {{Collection.Count}}
    /// </summary>
    private void ProcessCollectionMetadata(IXLRange range, Parameter[] pars)
    {
        // Find cells with collection metadata patterns
        var metadataCells = range.CellsUsed(cell =>
        {
            var value = cell.GetString();
            return value.Contains("{{") && value.Contains(".Count") && value.Contains("}}");
        });

        foreach (var cell in metadataCells)
        {
            string cellValue = cell.GetString();

            // Simple regex to find collection metadata expressions
            var matches = System.Text.RegularExpressions.Regex.Matches(
                cellValue, @"\{\{([^{}\.]+)\.Count\}\}");

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                string collectionName = match.Groups[1].Value;

                // Try to get the collection from variables
                if (_template.TryGetCollection(collectionName, out var collection))
                {
                    int count = collection.Count;
                    cellValue = cellValue.Replace(match.Value, count.ToString());
                }
            }

            // Update cell with processed value
            if (cellValue != cell.GetString())
            {
                cell.Value = cellValue;
            }
        }
    }
}