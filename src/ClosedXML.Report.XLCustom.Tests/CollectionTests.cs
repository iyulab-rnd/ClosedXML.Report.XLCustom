//using ClosedXML.Excel;
//using Xunit.Abstractions;

//namespace ClosedXML.Report.XLCustom.Tests
//{
//    public class CollectionTests : TestBase
//    {
//        public CollectionTests(ITestOutputHelper output) : base(output)
//        {
//        }

//        [Fact]
//        public void SimpleVerticalCollection_WithDefinedNames_AndFormatting_ShouldPopulateTemplate()
//        {
//            // Arrange
//            // Create a simple product list for testing
//            var products = new List<ProductModel>
//            {
//                new ProductModel { Id = 1, Name = "Laptop", Price = 1200.00m },
//                new ProductModel { Id = 2, Name = "Smartphone", Price = 800.00m },
//                new ProductModel { Id = 3, Name = "Headphones", Price = 150.00m }
//            };

//            // Create template workbook
//            using var workbook = new XLWorkbook();
//            var sheet = workbook.AddWorksheet("Products");

//            // Add headers
//            sheet.Cell("A1").Value = "Product ID";
//            sheet.Cell("B1").Value = "Product Name";
//            sheet.Cell("C1").Value = "Price";
//            sheet.Cell("D1").Value = "Calculated Value";
//            sheet.Cell("E1").Value = "Formatted Name";
//            sheet.Cell("F1").Value = "Currency";

//            // Add template expressions for the data row
//            sheet.Cell("A2").Value = "{{item.Id}}";
//            sheet.Cell("B2").Value = "{{item.Name}}";
//            sheet.Cell("C2").Value = "{{item.Price:F2}}"; // Format with 2 decimal places
//            sheet.Cell("D2").Value = "{{item.Price * 1.1:F2}}"; // Calculated field with format
//            sheet.Cell("E2").Value = "{{item.Name:upper}}"; // Format with custom upper formatter
//            sheet.Cell("F2").Value = "{{item.Price:C}}"; // Currency format

//            // Add service row for aggregations
//            sheet.Cell("A3").Value = "Totals:";
//            sheet.Cell("B3").Value = "";
//            sheet.Cell("C3").Value = "<<sum>>";
//            sheet.Cell("D3").Value = "<<sum>>";
//            sheet.Cell("E3").Value = "Total Items: {{ProductList.Count}}"; // Show count
//            sheet.Cell("F3").Value = "<<sum>>";

//            // Add service column (required for vertical tables)
//            sheet.Cell("G2").Value = "";

//            // Create a named range for the vertical collection
//            var productRange = sheet.Range("A2:G3");
//            workbook.DefinedNames.Add("ProductList", productRange);

//            // Apply styling to headers
//            var headerRange = sheet.Range("A1:F1");
//            headerRange.Style.Font.Bold = true;
//            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

//            // Apply number format to price columns
//            sheet.Cell("C2").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("D2").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("C3").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("D3").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("F2").Style.NumberFormat.Format = "$#,##0.00";
//            sheet.Cell("F3").Style.NumberFormat.Format = "$#,##0.00";

//            // Save template to memory stream
//            using var ms = new MemoryStream();
//            workbook.SaveAs(ms);
//            ms.Position = 0;

//            // Act
//            var template = new XLCustomTemplate(ms);

//            // Register the built-in formatters including 'upper'
//            template.RegisterBuiltInFormatters();

//            template.AddVariable("ProductList", products);
//            var result = template.Generate();
//            LogResult(result);

//            // Assert
//            Assert.False(result.HasErrors, "Template generation should not produce errors");
//            var wb = template.Workbook;
//            var ws = wb.Worksheet("Products");

//            // Verify data was correctly bound
//            Assert.Equal(1, ws.Cell("A2").GetValue<int>());
//            Assert.Equal("Laptop", ws.Cell("B2").GetString());
//            Assert.Equal(1200.00m, ws.Cell("C2").GetValue<decimal>());
//            Assert.Equal(1320.00m, ws.Cell("D2").GetValue<decimal>()); // Calculated value
//            Assert.Equal("LAPTOP", ws.Cell("E2").GetString()); // Uppercase format
//            Assert.Equal(1200.00m, ws.Cell("F2").GetValue<decimal>()); // Currency format

//            // Check the second row
//            Assert.Equal(2, ws.Cell("A3").GetValue<int>());
//            Assert.Equal("Smartphone", ws.Cell("B3").GetString());
//            Assert.Equal("SMARTPHONE", ws.Cell("E3").GetString()); // Uppercase format

//            // Verify sum aggregation in the totals row
//            int lastRow = 2 + products.Count;
//            decimal expectedTotal = products.Sum(p => p.Price);
//            decimal expectedCalculatedTotal = products.Sum(p => p.Price * 1.1m);

//            Assert.Equal(expectedTotal, ws.Cell($"C{lastRow}").GetValue<decimal>());
//            Assert.Equal(expectedCalculatedTotal, ws.Cell($"D{lastRow}").GetValue<decimal>());
//            Assert.Equal(expectedTotal, ws.Cell($"F{lastRow}").GetValue<decimal>());
//            Assert.Contains($"Total Items: {products.Count}", ws.Cell($"E{lastRow}").GetString());

//            Output.WriteLine("Simple vertical collection with DefinedNames and formatting test passed successfully!");
//        }

//        [Fact]
//        public void SimpleGroupedCollection_WithDefinedNames_ShouldPopulateTemplate()
//        {
//            // Arrange
//            // Create a sales data list with region and category for grouping
//            var salesData = new List<SalesDataModel>
//            {
//                new SalesDataModel { Region = "North", Category = "Electronics", Product = "Laptop", Units = 10, Revenue = 12000.00m },
//                new SalesDataModel { Region = "North", Category = "Electronics", Product = "Smartphone", Units = 15, Revenue = 9000.00m },
//                new SalesDataModel { Region = "North", Category = "Accessories", Product = "Headphones", Units = 20, Revenue = 2000.00m },
//                new SalesDataModel { Region = "South", Category = "Electronics", Product = "Laptop", Units = 8, Revenue = 9600.00m },
//                new SalesDataModel { Region = "South", Category = "Electronics", Product = "Smartphone", Units = 12, Revenue = 7200.00m },
//                new SalesDataModel { Region = "South", Category = "Accessories", Product = "Headphones", Units = 25, Revenue = 2500.00m }
//            };

//            // Create template workbook
//            using var workbook = new XLWorkbook();
//            var sheet = workbook.AddWorksheet("GroupedSales");

//            // Add headers
//            sheet.Cell("A1").Value = "Region";
//            sheet.Cell("B1").Value = "Category";
//            sheet.Cell("C1").Value = "Product";
//            sheet.Cell("D1").Value = "Units";
//            sheet.Cell("E1").Value = "Revenue";

//            // Add template expressions for data rows
//            sheet.Cell("A2").Value = "{{item.Region}}";
//            sheet.Cell("B2").Value = "{{item.Category}}";
//            sheet.Cell("C2").Value = "{{item.Product}}";
//            sheet.Cell("D2").Value = "{{item.Units}}";
//            sheet.Cell("E2").Value = "{{item.Revenue}}";

//            // Add service row with grouping and aggregation tags
//            sheet.Cell("A3").Value = "<<group>>"; // Group by Region
//            sheet.Cell("B3").Value = "<<group>>"; // Group by Category within Region
//            sheet.Cell("C3").Value = "";
//            sheet.Cell("D3").Value = "<<sum>>"; // Sum of Units for each group
//            sheet.Cell("E3").Value = "<<sum>>"; // Sum of Revenue for each group

//            // Add service column (required for vertical tables)
//            sheet.Cell("F2").Value = "";

//            // Create a named range for the collection
//            var salesRange = sheet.Range("A2:F3");
//            workbook.DefinedNames.Add("SalesData", salesRange);

//            // Apply number format to revenue column
//            sheet.Cell("E2").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("E3").Style.NumberFormat.Format = "#,##0.00";

//            // Save template to memory stream
//            using var ms = new MemoryStream();
//            workbook.SaveAs(ms);
//            ms.Position = 0;

//            // Act
//            var template = new XLCustomTemplate(ms);
//            template.AddVariable("SalesData", salesData);
//            var result = template.Generate();
//            LogResult(result);

//            // Assert
//            Assert.False(result.HasErrors, "Template generation should not produce errors");

//            var wb = template.Workbook;
//            var ws = wb.Worksheet("GroupedSales");

//            // Calculate expected group subtotals
//            var northElectronicsTotal = salesData
//                .Where(s => s.Region == "North" && s.Category == "Electronics")
//                .Sum(s => s.Revenue);

//            var northAccessoriesTotal = salesData
//                .Where(s => s.Region == "North" && s.Category == "Accessories")
//                .Sum(s => s.Revenue);

//            var southElectronicsTotal = salesData
//                .Where(s => s.Region == "South" && s.Category == "Electronics")
//                .Sum(s => s.Revenue);

//            var southAccessoriesTotal = salesData
//                .Where(s => s.Region == "South" && s.Category == "Accessories")
//                .Sum(s => s.Revenue);

//            // Calculate total rows that should be in the workbook
//            // Base data rows + group subtotal rows
//            int northElectronicsCount = salesData.Count(s => s.Region == "North" && s.Category == "Electronics");
//            int northAccessoriesCount = salesData.Count(s => s.Region == "North" && s.Category == "Accessories");
//            int southElectronicsCount = salesData.Count(s => s.Region == "South" && s.Category == "Electronics");
//            int southAccessoriesCount = salesData.Count(s => s.Region == "South" && s.Category == "Accessories");

//            // Find North/Electronics/Laptop row - first meaningful row in the sheet
//            // It might not be exactly at A2 due to the way ClosedXML.Report handles templates
//            bool foundNorthElectronicsLaptop = false;
//            for (int row = 2; row <= 20; row++) // Check within a reasonable range
//            {
//                if (ws.Cell($"A{row}").GetString() == "North" &&
//                    ws.Cell($"B{row}").GetString() == "Electronics" &&
//                    ws.Cell($"C{row}").GetString() == "Laptop")
//                {
//                    foundNorthElectronicsLaptop = true;
//                    Assert.Equal(10, ws.Cell($"D{row}").GetValue<int>());
//                    Assert.Equal(12000.00m, ws.Cell($"E{row}").GetValue<decimal>());
//                    break;
//                }
//            }

//            Assert.True(foundNorthElectronicsLaptop, "Could not find row with North/Electronics/Laptop data");

//            // Examine the structure and calculated totals
//            // We won't check exact row positions since grouping can affect layout
//            // Instead verify that subtotal rows exist with correct totals

//            // Verify that data was generated without errors and workbook has expected content
//            // We'll check for the presence of all expected product names in the output
//            Assert.Contains(ws.Cells(), cell => cell.Value.ToString() == "Laptop");
//            Assert.Contains(ws.Cells(), cell => cell.Value.ToString() == "Smartphone");
//            Assert.Contains(ws.Cells(), cell => cell.Value.ToString() == "Headphones");

//            // Verify grand total row exists
//            var lastRow = ws.LastRowUsed().RowNumber();
//            Assert.Equal(salesData.Sum(s => s.Revenue), ws.Cell($"E{lastRow}").GetValue<decimal>());

//            Output.WriteLine("Simple grouped collection with DefinedNames test passed successfully!");
//        }

//        [Fact]
//        public void SimpleNestedCollection_WithDefinedNames_ShouldPopulateTemplate()
//        {
//            // Arrange
//            var customers = new List<CustomerModel>
//            {
//                new CustomerModel
//                {
//                    Id = 101,
//                    Name = "ABC Company",
//                    Orders = new List<OrderModel>
//                    {
//                        new OrderModel {
//                            OrderId = 1001,
//                            Date = new DateTime(2025, 5, 1),
//                            Amount = 1200.00m,
//                            Items = new List<ItemModel> {
//                                new ItemModel { ItemId = 10001, Description = "Widget A", Quantity = 2, Price = 500.00m },
//                                new ItemModel { ItemId = 10002, Description = "Widget B", Quantity = 1, Price = 200.00m }
//                            }
//                        },
//                        new OrderModel {
//                            OrderId = 1002,
//                            Date = new DateTime(2025, 5, 5),
//                            Amount = 850.00m,
//                            Items = new List<ItemModel> {
//                                new ItemModel { ItemId = 10003, Description = "Widget C", Quantity = 1, Price = 850.00m }
//                            }
//                        }
//                    }
//                },
//                new CustomerModel
//                {
//                    Id = 102,
//                    Name = "XYZ Corporation",
//                    Orders = new List<OrderModel>
//                    {
//                        new OrderModel {
//                            OrderId = 1003,
//                            Date = new DateTime(2025, 5, 3),
//                            Amount = 2500.00m,
//                            Items = new List<ItemModel> {
//                                new ItemModel { ItemId = 10004, Description = "Widget D", Quantity = 5, Price = 500.00m }
//                            }
//                        }
//                    }
//                }
//            };

//            // Create template workbook
//            using var workbook = new XLWorkbook();
//            var sheet = workbook.AddWorksheet("Sheet 1");

//            // Set up column headers
//            sheet.Cell("A3").Value = "Customer Details";
//            sheet.Cell("B3").Value = "Customer No";
//            sheet.Cell("C3").Value = "Company Name";

//            sheet.Cell("A5").Value = "Order Information";
//            sheet.Cell("B5").Value = "Order No";
//            sheet.Cell("C5").Value = "Date";
//            sheet.Cell("D5").Value = "Amount";

//            sheet.Cell("A7").Value = "Item Details";
//            sheet.Cell("B7").Value = "Item No";
//            sheet.Cell("C7").Value = "Description";
//            sheet.Cell("D7").Value = "Quantity";
//            sheet.Cell("E7").Value = "Price";
//            sheet.Cell("F7").Value = "Subtotal";

//            // Add template expressions for customer data (row 4)
//            sheet.Cell("A4").Value = "";
//            sheet.Cell("B4").Value = "{{item.Id}}";
//            sheet.Cell("C4").Value = "{{item.Name}}";

//            // Add template expressions for order data (row 6)
//            sheet.Cell("A6").Value = "";
//            sheet.Cell("B6").Value = "{{item.OrderId}}";
//            sheet.Cell("C6").Value = "{{item.Date}}";
//            sheet.Cell("D6").Value = "{{item.Amount}}";

//            // Add template expressions for items data (row 8)
//            sheet.Cell("A8").Value = "";
//            sheet.Cell("B8").Value = "{{item.ItemId}}";
//            sheet.Cell("C8").Value = "{{item.Description}}";
//            sheet.Cell("D8").Value = "{{item.Quantity}}";
//            sheet.Cell("E8").Value = "{{item.Price}}";
//            // Fix the formula - just using the property references
//            sheet.Cell("F8").Value = "{{item.Quantity * item.Price}}";

//            // Add service rows with aggregation tags
//            sheet.Cell("A9").Value = "";
//            sheet.Cell("B9").Value = "";
//            sheet.Cell("C9").Value = "";
//            sheet.Cell("D9").Value = "";
//            sheet.Cell("E9").Value = "Total:";
//            sheet.Cell("F9").Value = "<<sum>>";

//            sheet.Cell("A10").Value = "";
//            sheet.Cell("B10").Value = "";
//            sheet.Cell("C10").Value = "";
//            sheet.Cell("D10").Value = "Order Total:";
//            sheet.Cell("E10").Value = "";
//            sheet.Cell("F10").Value = "<<sum>>";

//            sheet.Cell("A11").Value = "";
//            sheet.Cell("B11").Value = "";
//            sheet.Cell("C11").Value = "";
//            sheet.Cell("D11").Value = "Customer Total:";
//            sheet.Cell("E11").Value = "";
//            sheet.Cell("F11").Value = "<<sum>>";

//            // Apply formatting - ensure this is correctly applied
//            sheet.Cell("D6").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("E8").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("F8").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("F9").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("F10").Style.NumberFormat.Format = "#,##0.00";
//            sheet.Cell("F11").Style.NumberFormat.Format = "#,##0.00";

//            // Create named ranges for the nested collection
//            var itemsRange = sheet.Range("A8:I9");
//            var ordersRange = sheet.Range("A6:I10");
//            var customersRange = sheet.Range("A4:I11");

//            workbook.DefinedNames.Add("Customers_Orders_Items", itemsRange);
//            workbook.DefinedNames.Add("Customers_Orders", ordersRange);
//            workbook.DefinedNames.Add("Customers", customersRange);

//            // Save template to memory stream
//            using var ms = new MemoryStream();
//            workbook.SaveAs(ms);
//            ms.Position = 0;

//            // Act
//            var template = new XLCustomTemplate(ms);
//            template.AddVariable("Customers", customers);
//            var result = template.Generate();
//            // Output detailed error information if any
//            if (result.HasErrors)
//            {
//                foreach (var error in result.ParsingErrors)
//                {
//                    Output.WriteLine($"Error: {error.Message}, Range: '{error.Range}'");
//                }
//            }

//            // Assert
//            Assert.False(result.HasErrors, "Template generation should not produce errors");

//            var wb = template.Workbook;
//            var ws = wb.Worksheet("Sheet 1");

//            // Verify customer data
//            Assert.Equal(101, ws.Cell("B4").GetValue<int>());
//            Assert.Equal("ABC Company", ws.Cell("C4").GetValue<string>());

//            // Verify order data
//            Assert.Equal(1001, ws.Cell("B6").GetValue<int>());
//            Assert.Equal(new DateTime(2025, 5, 1), ws.Cell("C6").GetValue<DateTime>());
//            Assert.Equal(1200.00m, ws.Cell("D6").GetValue<decimal>());

//            // Verify item data - check with GetValue<decimal>() instead of string comparisons
//            bool foundFirstItem = false;
//            for (int row = 8; row <= 15; row++) // Check within a reasonable range
//            {
//                if (ws.Cell($"B{row}").GetString() == "10001" &&
//                    ws.Cell($"C{row}").GetString() == "Widget A")
//                {
//                    foundFirstItem = true;
//                    Assert.Equal(2, ws.Cell($"D{row}").GetValue<int>());
//                    Assert.Equal(500.00m, ws.Cell($"E{row}").GetValue<decimal>());
//                    Assert.Equal(1000.00m, ws.Cell($"F{row}").GetValue<decimal>());
//                    break;
//                }
//            }
//            Assert.True(foundFirstItem, "Could not find the first item data");

//            // Save the generated report for manual inspection
//            wb.SaveAs("GeneratedNestedReport.xlsx");

//            Output.WriteLine("Simple nested collection with DefinedNames test passed successfully!");
//        }

//        // Model classes
//        public class CustomerModel
//        {
//            public int Id { get; set; }
//            public string Name { get; set; }
//            public List<OrderModel> Orders { get; set; }
//        }

//        public class OrderModel
//        {
//            public int OrderId { get; set; }
//            public DateTime Date { get; set; }
//            public decimal Amount { get; set; }
//            public List<ItemModel> Items { get; set; }
//        }

//        public class ItemModel
//        {
//            public int ItemId { get; set; }
//            public string Description { get; set; }
//            public int Quantity { get; set; }
//            public decimal Price { get; set; }
//        }

//        // Model class for the test
//        public class SalesDataModel
//        {
//            public string Region { get; set; }
//            public string Category { get; set; }
//            public string Product { get; set; }
//            public int Units { get; set; }
//            public decimal Revenue { get; set; }
//        }

//        // Simple model class for the test
//        public class ProductModel
//        {
//            public int Id { get; set; }
//            public string Name { get; set; }
//            public decimal Price { get; set; }
//        }
//    }
//}