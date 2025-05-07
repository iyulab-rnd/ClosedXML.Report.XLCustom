using ClosedXML.Excel;
using FluentAssertions;
using System.Text;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests
{
    public class FormatTests : TestBase
    {
        public FormatTests(ITestOutputHelper output) : base(output)
        {
            // 각 테스트 실행 전에 싱글톤 리셋
            XLCustomRegistry.Reset();
        }

        [Fact]
        public void StandardNumericFormats_ShouldApplyCorrectly()
        {
            // Arrange
            var testModel = new FormatTestModel
            {
                Number = 1234.56m
            };

            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("NumericFormatTest");

            // 다양한 숫자 형식 테스트
            sheet.Cell("A1").Value = "Currency (C):";
            sheet.Cell("B1").Value = "{{Number:C}}";

            sheet.Cell("A2").Value = "Number (N2):";
            sheet.Cell("B2").Value = "{{Number:N2}}";

            sheet.Cell("A3").Value = "Percent (P):";
            sheet.Cell("B3").Value = "{{Number:P}}";

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;

            // Act
            var template = new XLCustomTemplate(ms).Preprocess();
            template.AddVariable(testModel);
            var result = template.Generate();

            // Assert - 실제 형식이 적용된 값을 검증
            result.HasErrors.Should().BeFalse();
            var ws = template.Workbook.Worksheet("NumericFormatTest");

            // 실제 보여지는 값도 검증
            var valueB1 = ws.Cell("B1").GetFormattedString();
            valueB1.Should().Contain("1,234.56");  // 통화 형식 ($ 기호는 시스템 설정에 따라 다를 수 있음)
        }

        public class FormatTestModel
        {
            public string Text { get; set; }
            public decimal Number { get; set; }
            public DateTime Date { get; set; }
        }
    }
}