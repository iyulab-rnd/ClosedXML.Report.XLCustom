Console.WriteLine("hello");

//using ClosedXML.Excel;
//using ClosedXML.Report.XLCustom;
//using System.Diagnostics;

//namespace ClosedXML.Report.XLCustom.TestConsoleApp
//{
//    internal class Program
//    {
//        static void Main(string[] args)
//        {
//            Console.WriteLine("ClosedXML.Report.XLCustom 테스트 콘솔 앱");
//            Console.WriteLine("======================================");

//            // 작업 디렉토리 설정 (현재 디렉토리)
//            string workingDir = AppDomain.CurrentDomain.BaseDirectory;
//            string templatePath = Path.Combine(workingDir, "ReportTemplate.xlsx");
//            string outputPath = Path.Combine(workingDir, "GeneratedReport.xlsx");

//            try
//            {
//                // 1. 샘플 데이터 생성
//                var data = CreateSampleData();

//                // 2. 템플릿 파일 생성
//                CreateTemplateFile(templatePath);
//                Console.WriteLine($"템플릿 파일이 생성되었습니다: {templatePath}");

//                // 3. 템플릿 파일을 사용하여 보고서 생성
//                GenerateReport(templatePath, outputPath, data);
//                Console.WriteLine($"생성된 보고서: {outputPath}");

//                // 4. 두 파일 모두 열기
//                OpenExcelFiles(templatePath, outputPath);
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine($"오류 발생: {ex.Message}");
//                Console.WriteLine(ex.StackTrace);
//            }

//            Console.WriteLine("\n프로그램 종료. 아무 키나 누르세요...");
//            Console.ReadKey();
//        }

//        static object CreateSampleData()
//        {
//            Console.WriteLine("샘플 데이터 생성 중...");

//            // 단순 속성
//            var data = new
//            {
//                // 기본 정보
//                Title = "분기별 판매 보고서",
//                ReportDate = DateTime.Now,
//                CompanyName = "ABC Corporation",
//                TotalRevenue = 1250000.50m,
//                IsApproved = true,

//                // 중첩 객체
//                Manager = new
//                {
//                    Name = "홍길동",
//                    Department = "영업부",
//                    Email = "hong@example.com",
//                    Phone = "010-1234-5678"
//                },

//                // 컬렉션
//                Products = new List<ProductItem>
//                {
//                    new ProductItem { Name = "제품 A", Price = 1200.50m, Sales = 145, InStock = true },
//                    new ProductItem { Name = "제품 B", Price = 250.75m, Sales = 267, InStock = true },
//                    new ProductItem { Name = "제품 C", Price = 7500.00m, Sales = 24, InStock = false },
//                    new ProductItem { Name = "제품 D", Price = 450.25m, Sales = 184, InStock = true },
//                    new ProductItem { Name = "제품 E", Price = 1800.99m, Sales = 97, InStock = true },
//                },

//                // 월별 판매 데이터
//                MonthlySales = new List<MonthlySalesData>
//                {
//                    new MonthlySalesData { Month = "1월", Revenue = 125000.50m, Expenses = 87500.25m },
//                    new MonthlySalesData { Month = "2월", Revenue = 135500.75m, Expenses = 92750.50m },
//                    new MonthlySalesData { Month = "3월", Revenue = 142800.25m, Expenses = 96250.75m },
//                }
//            };

//            return data;
//        }

//        static void CreateTemplateFile(string filePath)
//        {
//            Console.WriteLine("템플릿 파일 생성 중...");

//            using var workbook = new XLWorkbook();

//            // 기본 정보 워크시트
//            var infoSheet = workbook.AddWorksheet("기본 정보");

//            // 제목과 기본 정보
//            infoSheet.Cell("A1").Value = "{{Title}}";
//            infoSheet.Range("A1:D1").Merge();
//            infoSheet.Cell("A1").Style.Font.Bold = true;
//            infoSheet.Cell("A1").Style.Font.FontSize = 16;
//            infoSheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

//            // 일반 속성
//            infoSheet.Cell("A3").Value = "회사명:";
//            infoSheet.Cell("B3").Value = "{{CompanyName}}";

//            infoSheet.Cell("A4").Value = "보고서 날짜:";
//            infoSheet.Cell("B4").Value = "{{ReportDate:yyyy-MM-dd}}";  // Format 표현식

//            infoSheet.Cell("A5").Value = "총 수익:";
//            infoSheet.Cell("B5").Value = "{{TotalRevenue:C}}";  // 통화 형식

//            infoSheet.Cell("A6").Value = "승인 상태:";
//            infoSheet.Cell("B6").Value = "{{IsApproved}}";

//            // 관리자 정보 (중첩 객체)
//            infoSheet.Cell("A8").Value = "관리자 정보";
//            infoSheet.Cell("A8").Style.Font.Bold = true;

//            infoSheet.Cell("A9").Value = "이름:";
//            infoSheet.Cell("B9").Value = "{{Manager.Name}}";

//            infoSheet.Cell("A10").Value = "부서:";
//            infoSheet.Cell("B10").Value = "{{Manager.Department}}";

//            infoSheet.Cell("A11").Value = "이메일:";
//            infoSheet.Cell("B11").Value = "{{Manager.Email|link}}";  // 링크 함수

//            infoSheet.Cell("A12").Value = "전화번호:";
//            infoSheet.Cell("B12").Value = "{{Manager.Phone}}";

//            // 스타일 테스트
//            infoSheet.Cell("D3").Value = "포맷터 테스트";
//            infoSheet.Cell("D3").Style.Font.Bold = true;

//            infoSheet.Cell("D4").Value = "대문자:";
//            infoSheet.Cell("E4").Value = "{{CompanyName:upper}}";

//            infoSheet.Cell("D5").Value = "소문자:";
//            infoSheet.Cell("E5").Value = "{{CompanyName:lower}}";

//            infoSheet.Cell("D6").Value = "첫 글자 대문자:";
//            infoSheet.Cell("E6").Value = "{{CompanyName:titlecase}}";

//            // 함수 테스트
//            infoSheet.Cell("D8").Value = "함수 테스트";
//            infoSheet.Cell("D8").Style.Font.Bold = true;

//            infoSheet.Cell("D9").Value = "굵게:";
//            infoSheet.Cell("E9").Value = "{{CompanyName|bold}}";

//            infoSheet.Cell("D10").Value = "기울임:";
//            infoSheet.Cell("E10").Value = "{{CompanyName|italic}}";

//            infoSheet.Cell("D11").Value = "색상:";
//            infoSheet.Cell("E11").Value = "{{CompanyName|color(Red)}}";

//            // 제품 워크시트 (컬렉션 처리)
//            var productsSheet = workbook.AddWorksheet("제품 목록");

//            // 제목
//            productsSheet.Cell("A1").Value = "제품 목록";
//            productsSheet.Range("A1:F1").Merge();
//            productsSheet.Cell("A1").Style.Font.Bold = true;
//            productsSheet.Cell("A1").Style.Font.FontSize = 14;
//            productsSheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

//            // 헤더
//            productsSheet.Cell("A3").Value = "번호";
//            productsSheet.Cell("B3").Value = "제품명";
//            productsSheet.Cell("C3").Value = "가격";
//            productsSheet.Cell("D3").Value = "판매량";
//            productsSheet.Cell("E3").Value = "매출";
//            productsSheet.Cell("F3").Value = "재고";

//            // 헤더 스타일
//            var headerRange = productsSheet.Range("A3:F3");
//            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
//            headerRange.Style.Font.Bold = true;

//            // 데이터 행 템플릿
//            productsSheet.Cell("A4").Value = "{{i+1}}";
//            productsSheet.Cell("B4").Value = "{{item.Name|bold}}";  // 함수 적용
//            productsSheet.Cell("C4").Value = "{{item.Price:C}}";    // 형식 적용
//            productsSheet.Cell("D4").Value = "{{item.Sales}}";
//            productsSheet.Cell("E4").Value = "{{item.Price * item.Sales:C}}";  // 계산식 + 형식
//            productsSheet.Cell("F4").Value = "{{item.InStock}}";

//            // 서비스 행 (집계 함수)
//            productsSheet.Cell("A5").Value = "합계";
//            productsSheet.Cell("B5").Value = "";
//            productsSheet.Cell("C5").Value = "<<average>>";
//            productsSheet.Cell("D5").Value = "<<sum>>";
//            productsSheet.Cell("E5").Value = "<<sum over=\"item.Price * item.Sales\">>";
//            productsSheet.Cell("F5").Value = "";

//            // 서비스 열
//            productsSheet.Cell("G4").Value = "";

//            // 범위 정의
//            var productsRange = productsSheet.Range("A4:G5");
//            workbook.DefinedNames.Add("Products", productsRange);

//            // 월별 판매 워크시트
//            var salesSheet = workbook.AddWorksheet("월별 판매");

//            // 제목
//            salesSheet.Cell("A1").Value = "월별 판매 분석";
//            salesSheet.Range("A1:D1").Merge();
//            salesSheet.Cell("A1").Style.Font.Bold = true;
//            salesSheet.Cell("A1").Style.Font.FontSize = 14;
//            salesSheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

//            // 헤더
//            salesSheet.Cell("A3").Value = "월";
//            salesSheet.Cell("B3").Value = "수익";
//            salesSheet.Cell("C3").Value = "비용";
//            salesSheet.Cell("D3").Value = "순이익";

//            // 헤더 스타일
//            var salesHeaderRange = salesSheet.Range("A3:D3");
//            salesHeaderRange.Style.Fill.BackgroundColor = XLColor.LightGray;
//            salesHeaderRange.Style.Font.Bold = true;

//            // 데이터 행 템플릿
//            salesSheet.Cell("A4").Value = "{{item.Month}}";
//            salesSheet.Cell("B4").Value = "{{item.Revenue:C}}";
//            salesSheet.Cell("C4").Value = "{{item.Expenses:C}}";
//            salesSheet.Cell("D4").Value = "{{item.Revenue - item.Expenses:C}}";

//            // 조건부 서식 적용 (순이익이 양수면 녹색, 음수면 빨간색)
//            // 참고: 실제 엑셀 조건부 서식 대신 표현식으로 색상 적용
//            // 서비스 행 (집계 함수)
//            salesSheet.Cell("A5").Value = "합계";
//            salesSheet.Cell("B5").Value = "<<sum>>";
//            salesSheet.Cell("C5").Value = "<<sum>>";
//            salesSheet.Cell("D5").Value = "<<sum over=\"item.Revenue - item.Expenses\">>";

//            // 서비스 열
//            salesSheet.Cell("E4").Value = "";

//            // 범위 정의
//            var salesRange = salesSheet.Range("A4:E5");
//            workbook.DefinedNames.Add("MonthlySales", salesRange);

//            // 파일 저장
//            workbook.SaveAs(filePath);
//        }

//        static void GenerateReport(string templatePath, string outputPath, object data)
//        {
//            Console.WriteLine("보고서 생성 중...");

//            // 템플릿 생성
//            var template = new XLCustomTemplate(templatePath);

//            // 내장 포맷터 및 함수 등록
//            template.RegisterBuiltIns();

//            // Debug function registration
//            Console.WriteLine("Functions registered. Testing if they're available:");
//            Console.WriteLine($"Bold function registered: {template.GetRegisteredFunctions().Contains("bold")}");
//            Console.WriteLine($"Italic function registered: {template.GetRegisteredFunctions().Contains("italic")}");
//            Console.WriteLine($"Link function registered: {template.GetRegisteredFunctions().Contains("link")}");
//            Console.WriteLine($"Color function registered: {template.GetRegisteredFunctions().Contains("color")}");

//            // 변수 추가
//            template.AddVariable(data);

//            // 템플릿 생성
//            var result = template.Generate();

//            // Check for errors
//            if (result.HasErrors)
//            {
//                Console.WriteLine("Errors occurred during template processing:");
//                foreach (var error in result.ParsingErrors)
//                {
//                    Console.WriteLine($"- {error.Message} (Range: {error.Range})");
//                }
//            }
//            // Save result
//            template.SaveAs(outputPath);
//        }

//        static void OpenExcelFiles(string templatePath, string outputPath)
//        {
//            Console.WriteLine("파일 여는 중...");

//            // 템플릿 파일 열기
//            Process.Start(new ProcessStartInfo
//            {
//                FileName = templatePath,
//                UseShellExecute = true
//            });

//            // 잠시 대기
//            Thread.Sleep(1000);

//            // 생성된 보고서 파일 열기
//            Process.Start(new ProcessStartInfo
//            {
//                FileName = outputPath,
//                UseShellExecute = true
//            });
//        }
//    }

//    // 모델 클래스
//    public class ProductItem
//    {
//        public string Name { get; set; }
//        public decimal Price { get; set; }
//        public int Sales { get; set; }
//        public bool InStock { get; set; }
//    }

//    public class MonthlySalesData
//    {
//        public string Month { get; set; }
//        public decimal Revenue { get; set; }
//        public decimal Expenses { get; set; }
//    }
//}