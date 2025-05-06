using ClosedXML.Report.Options;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Tag registration functionality for XLCustomTemplate
/// </summary>
public partial class XLCustomTemplate
{
    /// <summary>
    /// Registers all standard tags with the template system
    /// </summary>
    private static void RegisterStandardTags()
    {
        // Ensure tags are registered only once
        lock (_registrationLock)
        {
            if (_tagsRegistered)
                return;

            // Range options
            TagsRegister.Add<RangeOptionTag>("Range", 255);
            TagsRegister.Add<RangeOptionTag>("SummaryAbove", 255);
            TagsRegister.Add<RangeOptionTag>("DisableGrandTotal", 255);

            // Grouping and pivot
            TagsRegister.Add<GroupTag>("Group", 200);
            TagsRegister.Add<PivotTag>("Pivot", 180);
            TagsRegister.Add<FieldPivotTag>("Row", 180);
            TagsRegister.Add<FieldPivotTag>("Column", 180);
            TagsRegister.Add<FieldPivotTag>("Col", 180);
            TagsRegister.Add<FieldPivotTag>("Page", 180);
            TagsRegister.Add<DataPivotTag>("Data", 180);

            // Sorting
            TagsRegister.Add<SortTag>("Sort", 128);
            TagsRegister.Add<SortTag>("Asc", 128);
            TagsRegister.Add<DescTag>("Desc", 128);

            // Summary functions
            TagsRegister.Add<SummaryFuncTag>("SUM", 50);
            TagsRegister.Add<SummaryFuncTag>("AVG", 50);
            TagsRegister.Add<SummaryFuncTag>("AVERAGE", 50);
            TagsRegister.Add<SummaryFuncTag>("COUNT", 50);
            TagsRegister.Add<SummaryFuncTag>("COUNTA", 50);
            TagsRegister.Add<SummaryFuncTag>("COUNTNUMS", 50);
            TagsRegister.Add<SummaryFuncTag>("MAX", 50);
            TagsRegister.Add<SummaryFuncTag>("MIN", 50);
            TagsRegister.Add<SummaryFuncTag>("PRODUCT", 50);
            TagsRegister.Add<SummaryFuncTag>("STDEV", 50);
            TagsRegister.Add<SummaryFuncTag>("STDEVP", 50);
            TagsRegister.Add<SummaryFuncTag>("VAR", 50);
            TagsRegister.Add<SummaryFuncTag>("VARP", 50);

            // Other formatting and behavior tags
            TagsRegister.Add<ImageTag>("Image", 100);
            TagsRegister.Add<OnlyValuesTag>("OnlyValues", 40);
            TagsRegister.Add<DeleteTag>("Delete", 5);
            TagsRegister.Add<AutoFilterTag>("AutoFilter", 0);
            TagsRegister.Add<ColsFitTag>("ColsFit", 0);
            TagsRegister.Add<RowsFitTag>("RowsFit", 0);
            TagsRegister.Add<HiddenTag>("Hidden", 0);
            TagsRegister.Add<HiddenTag>("Hide", 0);
            TagsRegister.Add<PageOptionsTag>("PageOptions", 0);
            TagsRegister.Add<ProtectedTag>("Protected", 0);
            TagsRegister.Add<HeightTag>("Height", 0);
            TagsRegister.Add<HeightRangeTag>("HeightRange", 0);
            TagsRegister.Add<ValidationTag>("Validation", 0);

            // Custom formatters and functions
            RegisterCustomFormatterTags();
            RegisterCustomFunctionTags();

            _tagsRegistered = true;
        }
    }

    /// <summary>
    /// Registers custom formatter tags with the template system
    /// </summary>
    private static void RegisterCustomFormatterTags()
    {
        // Add any custom formatter-specific tags here
    }

    /// <summary>
    /// Registers custom function tags with the template system
    /// </summary>
    private static void RegisterCustomFunctionTags()
    {
        // Add any custom function-specific tags here
    }

    // Static initialization tracking
    private static readonly object _registrationLock = new object();
    private static bool _tagsRegistered = false;
}