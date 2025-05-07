using ClosedXML.Report.Options;

namespace ClosedXML.Report.XLCustom
{
    public partial class XLCustomTemplate
    {
        static XLCustomTemplate()
        {
            RegisterDefaultTags();
        }

        // 기본 태그 등록 메서드
        private static void RegisterDefaultTags()
        {
            // XLTemplate 정적 생성자에서 등록하는 태그들을 동일하게 등록
            TagsRegister.Add<RangeOptionTag>("Range", 255);
            TagsRegister.Add<RangeOptionTag>("SummaryAbove", 255);
            TagsRegister.Add<RangeOptionTag>("DisableGrandTotal", 255);
            TagsRegister.Add<GroupTag>("Group", 200);
            TagsRegister.Add<PivotTag>("Pivot", 180);
            TagsRegister.Add<FieldPivotTag>("Row", 180);
            TagsRegister.Add<FieldPivotTag>("Column", 180);
            TagsRegister.Add<FieldPivotTag>("Col", 180);
            TagsRegister.Add<FieldPivotTag>("Page", 180);
            TagsRegister.Add<DataPivotTag>("Data", 180);
            TagsRegister.Add<SortTag>("Sort", 128);
            TagsRegister.Add<SortTag>("Asc", 128);
            TagsRegister.Add<DescTag>("Desc", 128);
            TagsRegister.Add<ImageTag>("Image", 100);
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
        }
    }
}