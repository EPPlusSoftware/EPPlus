using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Diagnostics;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class ReversedPolishNotationExecutionTests : TestBase
    {
        ExcelPackage _package;
        ParsingContext _parsingContext;
        RpnExpressionGraph _graph;
        private ISourceCodeTokenizer _tokenizer;
       [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            _parsingContext = ParsingContext.Create(_package);

            var dataProvider = new EpplusExcelDataProvider(_package, _parsingContext);
            _parsingContext.ExcelDataProvider = dataProvider;
            _parsingContext.NameValueProvider = new EpplusNameValueProvider(dataProvider);
            _parsingContext.RangeAddressFactory = new RangeAddressFactory(dataProvider, _parsingContext);

            _graph = new RpnExpressionGraph(_parsingContext);
            _tokenizer = OptimizedSourceCodeTokenizer.Default;
            
            
            SetUpWorksheet1();
            SetUpWorksheet2();
            _package.Workbook.Names.AddName("SumRange1", _package.Workbook.Worksheets["Sheet1"].Cells["A1:A100"]);
            _package.Workbook.Names.AddName("SumRange2", _package.Workbook.Worksheets["Sheet2"].Cells["A1:C2"]);
        }
        [TestCleanup]
        public void Cleanup()
        {
            SaveWorkbook("Rpn.xlsx", _package);

            _package.Dispose();
        }

        private void SetUpWorksheet1()
        {
            var ws1 = _package.Workbook.Worksheets.Add("Sheet1");
            ws1.Cells["A1:A1000"].Value = 1;
            ws1.Cells["B1"].Value = 100;
            ws1.Cells["B2:B1000"].Formula = "B1*(A2/A1)+1";

            ws1.Cells["C1"].Formula = "B1000/B1-1";
            ws1.Cells["D1"].Formula = "Sum(B1:B1000)/C1";
        }

        private void SetUpWorksheet2()
        {
            var ws2 = _package.Workbook.Worksheets.Add("Sheet2");
            ws2.Cells["A1"].Value = 1;
            ws2.Cells["B1"].Value = 2;
            ws2.Cells["C1"].Value = 3;

            ws2.Cells["A2"].Value = 10;
            ws2.Cells["B2"].Value = 20;
            ws2.Cells["C2"].Value = 30;

            ws2.Names.AddName("TwrStart", _package.Workbook.Worksheets["Sheet1"].Cells["B1"]);
            ws2.Cells["D1"].Formula = "Sum(SumRange1)+TwrStart";
            ws2.Cells["D2"].Formula = "Sum(SumRange2)+TwrStart";

            _package.Workbook.Names.AddValue("WorkbookDefinedNameValue", 1);
            ws2.Names.AddValue("WorksheetDefinedNameValue", "Name Value");
        }

        [TestMethod]
        public void ExecuteWorksheet1()
        {
            var dp = new RpnOptimizedDependencyChain(_package.Workbook, new ExcelCalculationOption());
            var sw=Stopwatch.StartNew();
            dp.Execute(_package.Workbook.Worksheets[0]);
            Debug.WriteLine($"Duration: {sw.ElapsedMilliseconds / 1000}");
            Assert.AreEqual(9.99D, _package.Workbook._worksheets[0].Cells["C1"].Value);
            Assert.AreEqual(60010.01D, Math.Round((double)_package.Workbook._worksheets[0].Cells["D1"].Value, 2));
        }
        [TestMethod]
        public void ExecuteWorksheet2()
        {
            var dp = new RpnOptimizedDependencyChain(_package.Workbook, new ExcelCalculationOption());
            dp.Execute(_package.Workbook.Worksheets[1]);

            Assert.AreEqual(200D, _package.Workbook._worksheets[1].Cells["D1"].Value);
            Assert.AreEqual(166D, _package.Workbook._worksheets[1].Cells["D2"].Value);
        }
        [TestMethod]
        public void CircularReferenceSelf()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "A1";
                ws.Cells["B1"].Formula = "A1:B1";

                var dc= RpnFormulaExecution.Execute(ws, new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(2, dc._circularReferences.Count);
                int wsIx, row, col;
                ExcelCellBase.SplitCellId(dc._circularReferences[0].FromCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 1);

                ExcelCellBase.SplitCellId(dc._circularReferences[0].ToCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 1);

                ExcelCellBase.SplitCellId(dc._circularReferences[1].FromCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 2);

                ExcelCellBase.SplitCellId(dc._circularReferences[1].ToCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 2);

            }
        }
        [TestMethod]
        public void CircularReferenceChain1()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "B1";
                ws.Cells["B1"].Formula = "C1";
                ws.Cells["C1"].Formula = "A1+1";

                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(1, dc._circularReferences.Count);
                int wsIx, row, col;
                ExcelCellBase.SplitCellId(dc._circularReferences[0].FromCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 3);

                ExcelCellBase.SplitCellId(dc._circularReferences[0].ToCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 1);
            }
        }
        [TestMethod]
        public void CircularReferenceChain2()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "B1";
                ws.Cells["B1"].Formula = "C1";
                ws.Cells["C1"].Formula = "Sum(A1:C1)";

                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(3, dc._circularReferences.Count);
                int wsIx, row, col;
                ExcelCellBase.SplitCellId(dc._circularReferences[0].FromCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 3);

                ExcelCellBase.SplitCellId(dc._circularReferences[0].ToCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 3);

                ExcelCellBase.SplitCellId(dc._circularReferences[1].FromCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 3);

                ExcelCellBase.SplitCellId(dc._circularReferences[1].ToCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 2);

                ExcelCellBase.SplitCellId(dc._circularReferences[2].FromCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 3);

                ExcelCellBase.SplitCellId(dc._circularReferences[2].ToCell, out wsIx, out row, out col);
                Assert.AreEqual(wsIx, 0);
                Assert.AreEqual(row, 1);
                Assert.AreEqual(col, 1);
            }
        }
        [TestMethod]
        public void IfFunctionTest1()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 2;
                ws.Cells["C1"].Formula = "if(A1 > B1, A1, Sum(b1))";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(2D, ws.Cells["C1"].Value);
            }
        }
        [TestMethod]
        public void IfFunctionTest2()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 2;
                ws.Cells["C1"].Formula = "if(A1 < B1, Offset(B1, 0, -1), Sum(b1))";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption() { AllowCircularReferences = true });
                Assert.AreEqual(1, ws.Cells["C1"].Value);
            }
        }
        [TestMethod]
        public void IfOffsetWithCircularReferenceTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 2;
                ws.Cells["C1"].Formula = "if(A1 < B1, Offset(C1, 0, -1), Sum(b1))";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                Assert.AreEqual(2, ws.Cells["C1"].Value);
            }
        }
        [TestMethod]
        public void IfOffsetInOffsetWithCircularReferenceTest()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 2;
                ws.Cells["C1"].Formula = "if(A1 < B1, Offset(Offset(B1, 0, A1), 0, -1), Sum(A1:B1))";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                Assert.AreEqual(2, ws.Cells["C1"].Value);
            }
        }
        [TestMethod]
        public void OffsetFunctionWithColonAddress()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 2;
                ws.Cells["C1"].Formula = "Sum(Offset(B1,0,-1):Offset(c1,0,-1))";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                Assert.AreEqual(3D, ws.Cells["C1"].Value);
            }
        }
        [TestMethod]
        public void IfFunctionWithColonAddress()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["B1"].Value = 2;
                ws.Cells["C1"].Formula = "Sum(if(true,B1,C1):If(false,C1,A1))";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                Assert.AreEqual(3D, ws.Cells["C1"].Value);
            }
        }
        [TestMethod]
        public void ColumnFunctionShouldNotReturnCircularReference()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["B15"].Formula = "Column(B1:B20)";
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                Assert.AreEqual(2, ws.Cells["B15"].Value);
            }
        }
        [TestMethod]
        public void SubtotalShouldNotIncludeSubtotalChildren_Avg()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Formula = "SUBTOTAL(1, A2:A3)";
                ws.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
                ws.Cells["A3"].Value = 2d;
                ws.Cells["A5"].Value = 2d;
                var dc = RpnFormulaExecution.Execute(ws, new ExcelCalculationOption());
                var result = ws.Cells["A1"].Value;
                Assert.AreEqual(2d, result);
            }
        }
    }
}
