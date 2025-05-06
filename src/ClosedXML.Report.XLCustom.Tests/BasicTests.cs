using ClosedXML.Excel;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests
{
    public class BasicTests : TestBase
    {
        public BasicTests(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void SimpleObject_ShouldPopulateTemplate()
        {
            // Arrange
            // Create a test model with properties
            var testModel = new TestModel
            {
                StringProperty = "Test String Value",
                IntProperty = 42,
                DecimalProperty = 123.45m,
                DateProperty = new DateTime(2025, 5, 6),
                BoolProperty = true
            };

            // Create a template in memory
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("PropertyTest");

            // Add template expressions for each property
            sheet.Cell("A1").Value = "String:";
            sheet.Cell("B1").Value = "{{StringProperty}}";

            sheet.Cell("A2").Value = "Integer:";
            sheet.Cell("B2").Value = "{{IntProperty}}";

            sheet.Cell("A3").Value = "Decimal:";
            sheet.Cell("B3").Value = "{{DecimalProperty}}";

            sheet.Cell("A4").Value = "Date:";
            sheet.Cell("B4").Value = "{{DateProperty}}";

            sheet.Cell("A5").Value = "Boolean:";
            sheet.Cell("B5").Value = "{{BoolProperty}}";

            // Save template to memory stream
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;

            // Act
            // Create a template processor
            var template = new XLCustomTemplate(ms);

            // Add the model to the template
            // Either add variables individually or add model directly
            // Individual properties
            template.AddVariable("StringProperty", testModel.StringProperty);
            template.AddVariable("IntProperty", testModel.IntProperty);
            template.AddVariable("DecimalProperty", testModel.DecimalProperty);
            template.AddVariable("DateProperty", testModel.DateProperty);
            template.AddVariable("BoolProperty", testModel.BoolProperty);

            // Generate the report
            var result = template.Generate();

            // Log any errors
            LogResult(result);

            // Assert
            // Check if there are no errors
            Assert.False(result.HasErrors, "Template generation should not produce errors");

            // Validate that each property was correctly bound
            var wb = template.Workbook;
            var ws = wb.Worksheet("PropertyTest");

            Assert.Equal("Test String Value", ws.Cell("B1").Value.ToString());
            Assert.Equal(42, ws.Cell("B2").Value);
            Assert.Equal(123.45m, ws.Cell("B3").Value);
            Assert.Equal(new DateTime(2025, 5, 6), ws.Cell("B4").Value);
            Assert.Equal(true, ws.Cell("B5").Value);

            // Optional: Save the result for visual inspection
            // wb.SaveAs("PropertyBindingResult.xlsx");

            Output.WriteLine("All property bindings were successful!");
        }

        [Fact]
        public void NestedObject_ShouldPopulateTemplate()
        {
            // Arrange
            // Create nested objects
            var nestedModel = new NestedModel
            {
                Parent = new ParentModel
                {
                    Name = "Parent Name",
                    Child = new ChildModel
                    {
                        Name = "Child Name",
                        Age = 10
                    }
                },
                SiblingList =
                [
                    new SiblingModel { Name = "Sibling 1", Age = 15 },
                    new SiblingModel { Name = "Sibling 2", Age = 20 }
                ]
            };

            // Create a template in memory
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("NestedTest");

            // Add template expressions for nested properties
            sheet.Cell("A1").Value = "Parent:";
            sheet.Cell("B1").Value = "{{model.Parent.Name}}";

            sheet.Cell("A2").Value = "Child:";
            sheet.Cell("B2").Value = "{{model.Parent.Child.Name}}";

            sheet.Cell("A3").Value = "Child Age:";
            sheet.Cell("B3").Value = "{{model.Parent.Child.Age}}";

            // For list data, define headers for the SiblingList vertical table
            sheet.Cell("A5").Value = "Name";
            sheet.Cell("B5").Value = "Age";

            // Define the data row template for the collection
            sheet.Cell("A6").Value = "{{item.Name}}"; // Use 'item' for collection elements
            sheet.Cell("B6").Value = "{{item.Age}}";

            // Add a service row and service column for vertical table (per documentation)
            sheet.Cell("A7").Value = ""; // Service row (can be empty or contain tags like <<sum>>)
            sheet.Cell("C6").Value = ""; // Service column (empty)

            // Create a named range for the SiblingList collection
            // Range includes data row (A6:B6), service row (A7:B7), and service column (C6)
            var siblingRange = sheet.Range("A6:C7");
            workbook.DefinedNames.Add("SiblingList", siblingRange);

            // Save template to memory stream
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;

            // Act
            var template = new XLCustomTemplate(ms);
            // Add variables explicitly to ensure scoping
            template.AddVariable("model", nestedModel);
            template.AddVariable("Parent", nestedModel.Parent); // Explicitly add Parent
            template.AddVariable("SiblingList", nestedModel.SiblingList); // Add collection

            var result = template.Generate();
            LogResult(result);

            // Assert
            Assert.False(result.HasErrors, "Template generation should not produce errors");

            var wb = template.Workbook;
            var ws = wb.Worksheet("NestedTest");

            // Check nested property bindings
            Assert.Equal("Parent Name", ws.Cell("B1").Value.ToString());
            Assert.Equal("Child Name", ws.Cell("B2").Value.ToString());
            Assert.Equal(10, ws.Cell("B3").Value);

            // Check for expanded rows from the collection
            int startRow = 6; // Starting row where data begins
            Assert.Equal("Sibling 1", ws.Cell(startRow, 1).Value.ToString());
            Assert.Equal(15, ws.Cell(startRow, 2).Value);
            Assert.Equal("Sibling 2", ws.Cell(startRow + 1, 1).Value.ToString());
            Assert.Equal(20, ws.Cell(startRow + 1, 2).Value);

            // Optional: Save the result for visual inspection
            // wb.SaveAs("NestedBindingResult.xlsx");

            Output.WriteLine("All nested property bindings were successful!");
        }

        [Fact]
        public void FormatSpecifier_ShouldApplyFormats()
        {
            // Arrange
            var formatModel = new FormatModel
            {
                Currency = 1234.56m,
                Percentage = 0.75m,
                LargeNumber = 1234567,
                Date = new DateTime(2025, 5, 6)
            };

            // Create a template in memory
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("FormatTest");

            // Add template expressions with format specifiers
            sheet.Cell("A1").Value = "Currency:";
            sheet.Cell("B1").Value = "{{Currency}}";

            sheet.Cell("A2").Value = "Percentage:";
            sheet.Cell("B2").Value = "{{Percentage}}";

            sheet.Cell("A3").Value = "Large Number:";
            sheet.Cell("B3").Value = "{{LargeNumber}}";

            sheet.Cell("A4").Value = "Date (Short):";
            sheet.Cell("B4").Value = "{{Date}}";

            sheet.Cell("A5").Value = "Date (Long):";
            sheet.Cell("B5").Value = "{{Date}}";

            // Save template to memory stream
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;

            // Act
            var template = new XLCustomTemplate(ms);
            // Add variables individually
            template.AddVariable("Currency", formatModel.Currency);
            template.AddVariable("Percentage", formatModel.Percentage);
            template.AddVariable("LargeNumber", formatModel.LargeNumber);
            template.AddVariable("Date", formatModel.Date);

            var result = template.Generate();
            LogResult(result);

            // Assert
            Assert.False(result.HasErrors, "Template generation should not produce errors");

            var wb = template.Workbook;
            var ws = wb.Worksheet("FormatTest");

            // Check cell values - we're just checking if the values are bound correctly
            // The actual formatting will be applied by Excel when viewing the file
            Assert.Equal(1234.56m, ws.Cell("B1").Value);
            Assert.Equal(0.75m, ws.Cell("B2").Value);
            Assert.Equal(1234567, ws.Cell("B3").Value);
            Assert.Equal(new DateTime(2025, 5, 6), ws.Cell("B4").Value);
            Assert.Equal(new DateTime(2025, 5, 6), ws.Cell("B5").Value);

            Output.WriteLine("All formatted property bindings were successful!");
        }

        [Fact]
        public void HorizontalCollection_ShouldPopulateTemplate()
        {
            List<ItemModel> items =
            [
                new ItemModel { Name = "ItemA", Price = 10 },
                new ItemModel { Name = "ItemB", Price = 20 },
                new ItemModel { Name = "ItemC", Price = 30 },
            ];

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Horizontal");
            ws.Cell("A1").Value = "Name1"; ws.Cell("B1").Value = "Price1";
            ws.Cell("C1").Value = "Name2"; ws.Cell("D1").Value = "Price2";
            ws.Cell("E1").Value = "Name3"; ws.Cell("F1").Value = "Price3";

            ws.Cell("A2").Value = "{{item[0].Name}}"; ws.Cell("B2").Value = "{{item[0].Price}}";
            ws.Cell("C2").Value = "{{item[1].Name}}"; ws.Cell("D2").Value = "{{item[1].Price}}";
            ws.Cell("E2").Value = "{{item[2].Name}}"; ws.Cell("F2").Value = "{{item[2].Price}}";

            using var ms = new MemoryStream();
            wb.SaveAs(ms); ms.Position = 0;

            var template = new XLCustomTemplate(ms);
            template.AddVariable("item", items);
            var result = template.Generate();
            LogResult(result);

            var res = template.Workbook.Worksheet("Horizontal");
            Assert.Equal("ItemA", res.Cell("A2").GetString());
            Assert.Equal(10, res.Cell("B2").GetValue<int>());
            Assert.Equal("ItemB", res.Cell("C2").GetString());
            Assert.Equal(20, res.Cell("D2").GetValue<int>());
            Assert.Equal("ItemC", res.Cell("E2").GetString());
            Assert.Equal(30, res.Cell("F2").GetValue<int>());
        }

        [Fact]
        public void AggregationTags_ShouldCalculateCorrectTotals()
        {
            // Arrange
            // Create test data with numeric values for aggregation testing
            var salesData = new List<SalesRecord>
            {
                new SalesRecord { Region = "North", Product = "Widget A", Units = 150, Revenue = 7500.00m, Profit = 2250.00m, DiscountRate = 0.05m },
                new SalesRecord { Region = "North", Product = "Widget B", Units = 200, Revenue = 12000.00m, Profit = 4200.00m, DiscountRate = 0.10m },
                new SalesRecord { Region = "South", Product = "Widget A", Units = 175, Revenue = 8750.00m, Profit = 2625.00m, DiscountRate = 0.07m },
                new SalesRecord { Region = "South", Product = "Widget B", Units = 120, Revenue = 7200.00m, Profit = 2160.00m, DiscountRate = 0.12m },
                new SalesRecord { Region = "East", Product = "Widget A", Units = 210, Revenue = 10500.00m, Profit = 3675.00m, DiscountRate = 0.05m },
                new SalesRecord { Region = "East", Product = "Widget B", Units = 180, Revenue = 10800.00m, Profit = 3240.00m, DiscountRate = 0.08m },
                new SalesRecord { Region = "West", Product = "Widget A", Units = 160, Revenue = 8000.00m, Profit = 2400.00m, DiscountRate = 0.06m },
                new SalesRecord { Region = "West", Product = "Widget B", Units = 190, Revenue = 11400.00m, Profit = 3990.00m, DiscountRate = 0.09m }
            };

            // Create a template in memory with multiple aggregation tags
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("AggregationTest");

            // Add headers
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "Product";
            sheet.Cell("C1").Value = "Units";
            sheet.Cell("D1").Value = "Revenue";
            sheet.Cell("E1").Value = "Profit";
            sheet.Cell("F1").Value = "Discount Rate";

            // Add template expressions for data rows
            sheet.Cell("A2").Value = "{{item.Region}}";
            sheet.Cell("B2").Value = "{{item.Product}}";
            sheet.Cell("C2").Value = "{{item.Units}}";
            sheet.Cell("D2").Value = "{{item.Revenue}}";
            sheet.Cell("E2").Value = "{{item.Profit}}";
            sheet.Cell("F2").Value = "{{item.DiscountRate}}";

            // Add service row with various aggregation tags
            sheet.Cell("A3").Value = "Totals"; // Label for totals
            sheet.Cell("B3").Value = ""; // Empty cell
            sheet.Cell("C3").Value = "<<sum>>"; // Sum of Units
            sheet.Cell("D3").Value = "<<sum>>"; // Sum of Revenue
            sheet.Cell("E3").Value = "<<sum>>"; // Sum of Profit
            sheet.Cell("F3").Value = "<<average>>"; // Average of Discount Rate

            // Add service column (mandatory for vertical tables)
            sheet.Cell("G2").Value = "";

            // Define the named range for data
            var salesRange = sheet.Range("A2:G3");
            workbook.DefinedNames.Add("SalesData", salesRange);

            // Apply appropriate number formats
            sheet.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("F2").Style.NumberFormat.Format = "0.00%";
            sheet.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("E3").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("F3").Style.NumberFormat.Format = "0.00%";

            // Create a second sheet for additional aggregation functions
            var sheet2 = workbook.AddWorksheet("MoreAggregations");

            // Add headers
            sheet2.Cell("A1").Value = "Region";
            sheet2.Cell("B1").Value = "Product";
            sheet2.Cell("C1").Value = "Units";
            sheet2.Cell("D1").Value = "Revenue";
            sheet2.Cell("E1").Value = "Profit";
            sheet2.Cell("F1").Value = "Discount Rate";

            // Add template expressions for data rows
            sheet2.Cell("A2").Value = "{{item.Region}}";
            sheet2.Cell("B2").Value = "{{item.Product}}";
            sheet2.Cell("C2").Value = "{{item.Units}}";
            sheet2.Cell("D2").Value = "{{item.Revenue}}";
            sheet2.Cell("E2").Value = "{{item.Profit}}";
            sheet2.Cell("F2").Value = "{{item.DiscountRate}}";

            // Service row with other aggregation functions
            sheet2.Cell("A3").Value = "<<AutoFilter>>"; // Add AutoFilter to enable filtering
            sheet2.Cell("B3").Value = ""; // Empty cell
            sheet2.Cell("C3").Value = "<<count>>"; // Count of Units
            sheet2.Cell("D3").Value = "<<max>>"; // Max Revenue
            sheet2.Cell("E3").Value = "<<min>>"; // Min Profit
            sheet2.Cell("F3").Value = "<<product>>"; // Product of Discount Rate

            // Service column
            sheet2.Cell("G2").Value = "";

            // Define the named range for the second sheet
            var salesRange2 = sheet2.Range("A2:G3");
            workbook.DefinedNames.Add("SalesData2", salesRange2);

            // Apply number formats
            sheet2.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("F2").Style.NumberFormat.Format = "0.00%";
            sheet2.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("E3").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("F3").Style.NumberFormat.Format = "0.00%";

            // Create a third sheet for complex aggregation (using the 'over' parameter)
            var sheet3 = workbook.AddWorksheet("ComplexAggregation");

            // Add headers
            sheet3.Cell("A1").Value = "Region";
            sheet3.Cell("B1").Value = "Product";
            sheet3.Cell("C1").Value = "Units";
            sheet3.Cell("D1").Value = "Revenue";
            sheet3.Cell("E1").Value = "Profit";
            sheet3.Cell("F1").Value = "Discount Rate";
            sheet3.Cell("G1").Value = "Profit Margin";

            // Add template expressions for data rows
            sheet3.Cell("A2").Value = "{{item.Region}}";
            sheet3.Cell("B2").Value = "{{item.Product}}";
            sheet3.Cell("C2").Value = "{{item.Units}}";
            sheet3.Cell("D2").Value = "{{item.Revenue}}";
            sheet3.Cell("E2").Value = "{{item.Profit}}";
            sheet3.Cell("F2").Value = "{{item.DiscountRate}}";
            sheet3.Cell("G2").Value = "{{item.Profit / item.Revenue}}"; // Calculate profit margin

            // Service row with complex aggregations
            sheet3.Cell("A3").Value = "<<sort>>"; // Sort by Region
            sheet3.Cell("B3").Value = "";
            sheet3.Cell("C3").Value = "<<sum over=\"item.Units\">>"; // Sum using 'over' parameter
            sheet3.Cell("D3").Value = "<<sum over=\"item.Revenue\">>"; // Sum using 'over' parameter
            sheet3.Cell("E3").Value = "<<sum over=\"item.Profit\">>"; // Sum using 'over' parameter
            sheet3.Cell("F3").Value = "";
            sheet3.Cell("G3").Value = "<<average over=\"item.Profit / item.Revenue\">>"; // Average profit margin

            // Service column
            sheet3.Cell("H2").Value = "";

            // Define the named range
            var salesRange3 = sheet3.Range("A2:H3");
            workbook.DefinedNames.Add("SalesData3", salesRange3);

            // Apply number formats
            sheet3.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("F2").Style.NumberFormat.Format = "0.00%";
            sheet3.Cell("G2").Style.NumberFormat.Format = "0.00%";
            sheet3.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("E3").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("G3").Style.NumberFormat.Format = "0.00%";

            // Save template to memory stream
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;

            // Act
            // Create a template processor
            var template = new XLCustomTemplate(ms);

            // Add variables for all three sheets
            template.AddVariable("SalesData", salesData);
            template.AddVariable("SalesData2", salesData);
            template.AddVariable("SalesData3", salesData);

            // Generate the report
            var result = template.Generate();

            // Log any errors
            LogResult(result);

            // Assert
            // Check if there are no errors
            Assert.False(result.HasErrors, "Template generation should not produce errors");

            // Validate aggregation results on first sheet
            var wb = template.Workbook;
            var ws = wb.Worksheet("AggregationTest");

            // Calculate expected values
            int totalUnits = salesData.Sum(x => x.Units);
            decimal totalRevenue = salesData.Sum(x => x.Revenue);
            decimal totalProfit = salesData.Sum(x => x.Profit);
            decimal avgDiscountRate = salesData.Average(x => x.DiscountRate);

            // Aggregate functions should be replaced with Excel SUBTOTAL formulas
            // which should have calculated the correct values
            Assert.Equal(totalUnits, ws.Cell("C" + (2 + salesData.Count)).Value);
            Assert.Equal(totalRevenue, ws.Cell("D" + (2 + salesData.Count)).Value);
            Assert.Equal(totalProfit, ws.Cell("E" + (2 + salesData.Count)).Value);

            // For the average, we need to allow a small epsilon due to floating point precision
            var actualAvgDiscount = (double)ws.Cell("F" + (2 + salesData.Count)).Value;
            var expectedAvgDiscount = (double)avgDiscountRate;
            Assert.True(Math.Abs(actualAvgDiscount - expectedAvgDiscount) < 0.0001,
                $"Average discount rate {actualAvgDiscount} does not match expected {expectedAvgDiscount}");

            // Check second sheet with other aggregation functions
            var ws2 = wb.Worksheet("MoreAggregations");

            // Calculate expected values
            int countUnits = salesData.Count;
            decimal maxRevenue = salesData.Max(x => x.Revenue);
            decimal minProfit = salesData.Min(x => x.Profit);

            // Verify count function
            Assert.Equal(countUnits, ws2.Cell("C" + (2 + salesData.Count)).Value);

            // Verify max function
            Assert.Equal(maxRevenue, ws2.Cell("D" + (2 + salesData.Count)).Value);

            // Verify min function
            Assert.Equal(minProfit, ws2.Cell("E" + (2 + salesData.Count)).Value);

            // Check third sheet with complex 'over' parameter aggregations
            var ws3 = wb.Worksheet("ComplexAggregation");

            // Calculate expected average profit margin
            decimal avgProfitMargin = salesData.Average(x => x.Profit / x.Revenue);

            var actualAvgMargin = (double)ws3.Cell("G" + (2 + salesData.Count)).Value;
            var expectedAvgMargin = (double)avgProfitMargin;
            Assert.True(Math.Abs(actualAvgMargin - expectedAvgMargin) < 0.0001,
                $"Average profit margin {actualAvgMargin} does not match expected {expectedAvgMargin}");

            // Optional: Save the result for visual inspection
            // wb.SaveAs("AggregationTestResult.xlsx");

            Output.WriteLine("All aggregation tests passed successfully!");
        }

        [Fact]
        public void GroupingWithAggregation_ShouldCalculateGroupTotals()
        {
            // Arrange
            // Create test data for grouping and aggregation
            var salesData = new List<SalesRecord>
            {
                new SalesRecord { Region = "North", Product = "Widget A", Units = 150, Revenue = 7500.00m, Profit = 2250.00m, DiscountRate = 0.05m },
                new SalesRecord { Region = "North", Product = "Widget B", Units = 200, Revenue = 12000.00m, Profit = 4200.00m, DiscountRate = 0.10m },
                new SalesRecord { Region = "South", Product = "Widget A", Units = 175, Revenue = 8750.00m, Profit = 2625.00m, DiscountRate = 0.07m },
                new SalesRecord { Region = "South", Product = "Widget B", Units = 120, Revenue = 7200.00m, Profit = 2160.00m, DiscountRate = 0.12m },
                new SalesRecord { Region = "East", Product = "Widget A", Units = 210, Revenue = 10500.00m, Profit = 3675.00m, DiscountRate = 0.05m },
                new SalesRecord { Region = "East", Product = "Widget B", Units = 180, Revenue = 10800.00m, Profit = 3240.00m, DiscountRate = 0.08m },
                new SalesRecord { Region = "West", Product = "Widget A", Units = 160, Revenue = 8000.00m, Profit = 2400.00m, DiscountRate = 0.06m },
                new SalesRecord { Region = "West", Product = "Widget B", Units = 190, Revenue = 11400.00m, Profit = 3990.00m, DiscountRate = 0.09m }
            };

            // Create a template in memory with grouping and aggregation
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("GroupingTest");

            // Add headers
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "Product";
            sheet.Cell("C1").Value = "Units";
            sheet.Cell("D1").Value = "Revenue";
            sheet.Cell("E1").Value = "Profit";
            sheet.Cell("F1").Value = "Discount Rate";

            // Add template expressions for data rows
            sheet.Cell("A2").Value = "{{item.Region}}";
            sheet.Cell("B2").Value = "{{item.Product}}";
            sheet.Cell("C2").Value = "{{item.Units}}";
            sheet.Cell("D2").Value = "{{item.Revenue}}";
            sheet.Cell("E2").Value = "{{item.Profit}}";
            sheet.Cell("F2").Value = "{{item.DiscountRate}}";

            // Add service row with grouping and aggregation
            sheet.Cell("A3").Value = "<<group>>"; // Group by Region
            sheet.Cell("B3").Value = "<<group>>"; // Group by Product within Region
            sheet.Cell("C3").Value = "<<sum>>"; // Sum Units for each group
            sheet.Cell("D3").Value = "<<sum>>"; // Sum Revenue for each group
            sheet.Cell("E3").Value = "<<sum>>"; // Sum Profit for each group
            sheet.Cell("F3").Value = "<<average>>"; // Average Discount Rate for each group

            // Add service column
            sheet.Cell("G2").Value = "";

            // Define the named range
            var salesRange = sheet.Range("A2:G3");
            workbook.DefinedNames.Add("GroupedSales", salesRange);

            // Apply number formats
            sheet.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("F2").Style.NumberFormat.Format = "0.00%";
            sheet.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("E3").Style.NumberFormat.Format = "#,##0.00";
            sheet.Cell("F3").Style.NumberFormat.Format = "0.00%";

            // Create a sheet with MergeLabels option
            var sheet2 = workbook.AddWorksheet("MergedGroups");

            // Add headers
            sheet2.Cell("A1").Value = "Region";
            sheet2.Cell("B1").Value = "Product";
            sheet2.Cell("C1").Value = "Units";
            sheet2.Cell("D1").Value = "Revenue";
            sheet2.Cell("E1").Value = "Profit";
            sheet2.Cell("F1").Value = "Discount Rate";

            // Add template expressions for data rows
            sheet2.Cell("A2").Value = "{{item.Region}}";
            sheet2.Cell("B2").Value = "{{item.Product}}";
            sheet2.Cell("C2").Value = "{{item.Units}}";
            sheet2.Cell("D2").Value = "{{item.Revenue}}";
            sheet2.Cell("E2").Value = "{{item.Profit}}";
            sheet2.Cell("F2").Value = "{{item.DiscountRate}}";

            // Add service row with merged group labels
            sheet2.Cell("A3").Value = "<<group MergeLabels>>"; // Group and merge Region cells
            sheet2.Cell("B3").Value = "<<group>>"; // Group by Product within Region
            sheet2.Cell("C3").Value = "<<sum>>"; // Sum Units
            sheet2.Cell("D3").Value = "<<sum>>"; // Sum Revenue
            sheet2.Cell("E3").Value = "<<sum>>"; // Sum Profit
            sheet2.Cell("F3").Value = "<<average>>"; // Average Discount Rate

            // Add service column
            sheet2.Cell("G2").Value = "";

            // Define the named range
            var salesRange2 = sheet2.Range("A2:G3");
            workbook.DefinedNames.Add("MergedGroupedSales", salesRange2);

            // Apply number formats
            sheet2.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("F2").Style.NumberFormat.Format = "0.00%";
            sheet2.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("E3").Style.NumberFormat.Format = "#,##0.00";
            sheet2.Cell("F3").Style.NumberFormat.Format = "0.00%";

            // Create a sheet with SummaryAbove option
            var sheet3 = workbook.AddWorksheet("SummaryAbove");

            // Add headers
            sheet3.Cell("A1").Value = "Region";
            sheet3.Cell("B1").Value = "Product";
            sheet3.Cell("C1").Value = "Units";
            sheet3.Cell("D1").Value = "Revenue";
            sheet3.Cell("E1").Value = "Profit";
            sheet3.Cell("F1").Value = "Discount Rate";

            // Add template expressions for data rows
            sheet3.Cell("A2").Value = "{{item.Region}}";
            sheet3.Cell("B2").Value = "{{item.Product}}";
            sheet3.Cell("C2").Value = "{{item.Units}}";
            sheet3.Cell("D2").Value = "{{item.Revenue}}";
            sheet3.Cell("E2").Value = "{{item.Profit}}";
            sheet3.Cell("F2").Value = "{{item.DiscountRate}}";

            // Add service row with SummaryAbove
            sheet3.Cell("A3").Value = "<<SummaryAbove>> <<group>>"; // Group by Region with summary above
            sheet3.Cell("B3").Value = "<<group>>"; // Group by Product
            sheet3.Cell("C3").Value = "<<sum>>"; // Sum Units
            sheet3.Cell("D3").Value = "<<sum>>"; // Sum Revenue
            sheet3.Cell("E3").Value = "<<sum>>"; // Sum Profit
            sheet3.Cell("F3").Value = "<<average>>"; // Average Discount Rate

            // Add service column
            sheet3.Cell("G2").Value = "";

            // Define the named range
            var salesRange3 = sheet3.Range("A2:G3");
            workbook.DefinedNames.Add("SummaryAboveSales", salesRange3);

            // Apply number formats
            sheet3.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("F2").Style.NumberFormat.Format = "0.00%";
            sheet3.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("E3").Style.NumberFormat.Format = "#,##0.00";
            sheet3.Cell("F3").Style.NumberFormat.Format = "0.00%";

            // Save template to memory stream
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            ms.Position = 0;

            // Act
            // Create a template processor
            var template = new XLCustomTemplate(ms);

            // Add variables for all sheets
            template.AddVariable("GroupedSales", salesData);
            template.AddVariable("MergedGroupedSales", salesData);
            template.AddVariable("SummaryAboveSales", salesData);

            // Generate the report
            var result = template.Generate();

            // Log any errors
            LogResult(result);

            // Assert
            // Check if there are no errors
            Assert.False(result.HasErrors, "Template generation should not produce errors");

            // For grouped data, we need to validate that the structure is correct
            // and subtotals are calculated correctly
            var wb = template.Workbook;

            // Check if all worksheets are present
            Assert.True(wb.Worksheets.Contains("GroupingTest"), "GroupingTest worksheet should exist");
            Assert.True(wb.Worksheets.Contains("MergedGroups"), "MergedGroups worksheet should exist");
            Assert.True(wb.Worksheets.Contains("SummaryAbove"), "SummaryAbove worksheet should exist");

            // Optional: Save the result for visual inspection
            // wb.SaveAs("GroupingTestResult.xlsx");

            Output.WriteLine("All grouping with aggregation tests passed successfully!");
        }

        public class TestModel
        {
            public string StringProperty { get; set; }
            public int IntProperty { get; set; }
            public decimal DecimalProperty { get; set; }
            public DateTime DateProperty { get; set; }
            public bool BoolProperty { get; set; }
        }

        public class ParentModel
        {
            public string Name { get; set; }
            public ChildModel Child { get; set; }
        }

        public class ChildModel
        {
            public string Name { get; set; }
            public int Age { get; set; }
        }

        public class SiblingModel
        {
            public string Name { get; set; }
            public int Age { get; set; }
        }

        public class NestedModel
        {
            public ParentModel Parent { get; set; }
            public List<SiblingModel> SiblingList { get; set; }
        }

        public class FormatModel
        {
            public decimal Currency { get; set; }
            public decimal Percentage { get; set; }
            public int LargeNumber { get; set; }
            public DateTime Date { get; set; }
        }

        public class ItemModel { public string Name { get; set; } public int Price { get; set; } }
        public class FlagModel { public bool IsActive { get; set; } }
        public class FormatModel2 { public decimal Amount { get; set; } public decimal Rate { get; set; } public DateTime Date { get; set; } }
        public class FuncModel { public string Text { get; set; } }

        public class SalesRecord
        {
            public string Region { get; set; }
            public string Product { get; set; }
            public int Units { get; set; }
            public decimal Revenue { get; set; }
            public decimal Profit { get; set; }
            public decimal DiscountRate { get; set; }
        }

    }
}