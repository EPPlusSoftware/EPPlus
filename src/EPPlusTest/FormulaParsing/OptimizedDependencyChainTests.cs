using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class OptimizedDependencyChainTests : TestBase
    {
        static ExcelPackage _package;
        static EpplusExcelDataProvider _excelDataProvider;
        static ExpressionGraphBuilder _graphBuilder;
        static ExcelWorksheet _ws;
        internal static ISourceCodeTokenizer _tokenizer = OptimizedSourceCodeTokenizer.Default;
        static ExpressionCompiler _compiler;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _package = new ExcelPackage();

            _ws = _package.Workbook.Worksheets.Add("Sheet1");
            LoadTestdata(_ws);
            var tbl = _ws.Tables.Add(_ws.Cells["A1:E101"], "MyTable");
            tbl.ShowTotal = true;
            _excelDataProvider = new EpplusExcelDataProvider(_package);
            var parsingContext = ParsingContext.Create(_package);
            _compiler = new ExpressionCompiler(parsingContext);

            parsingContext.CurrentCell = new FormulaCellAddress(0, 1, 1);
            parsingContext.ExcelDataProvider = _excelDataProvider;
            _graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, parsingContext);
        }
        [ClassCleanup]
        public void Cleanup()
        {
            SaveWorkbook("DependencyChain.xlsx", _package);
            _package.Dispose();
        }
        [TestMethod]
        public void VerifyCellAddressExpression_NonFixed()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {                
                var ws = p.Workbook.Worksheets[0];
                var dp=OptimizedDependencyChainFactory.Create(p.Workbook, new ExcelCalculationOption(){ });
            }
        }
        [TestMethod]
        public void VerifyRangeAddressExpression_Range()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var dp = OptimizedDependencyChainFactory.Create(ws, new ExcelCalculationOption() { });
            }
        }
        [TestMethod]
        public void VerifyTableAddressExpression_Table()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[2];
                var dp = OptimizedDependencyChainFactory.Create(ws, new ExcelCalculationOption() { });
            }
        }
        [TestMethod]
        public void VerifyAddressExpressions_CrossReference()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[3];
                var dp = OptimizedDependencyChainFactory.Create(ws, new ExcelCalculationOption() { });
            }
        }
        [TestMethod]
        public void VerifyAddressExpressions_CirularReference1()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[4];
                var dp = OptimizedDependencyChainFactory.Create(ws.Cells["A1"], new ExcelCalculationOption() { });
            }
        }
        [TestMethod]
        public void VerifyAddressExpressions_CirularReference2()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[4];
                var dp = OptimizedDependencyChainFactory.Create(ws.Cells["A2"], new ExcelCalculationOption() { });
                Assert.AreEqual(1, dp._circularReferences.Count);
            }
        }
        [TestMethod]
        public void VerifyAddressExpressions_CirularReference3()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[4];
                var dp = OptimizedDependencyChainFactory.Create(ws.Cells["A3"], new ExcelCalculationOption() { });
                Assert.AreEqual(1, dp._circularReferences.Count);
            }
        }
        [TestMethod]
        public void VerifyDefinedNames()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[5];
                var dp = OptimizedDependencyChainFactory.Create(ws.Cells["A1:A10"], new ExcelCalculationOption() { });
                Assert.AreEqual(1, dp._circularReferences.Count);
            }
        }
        [TestMethod]
        public void VerifyIfFunction()
        {
            using (var p = OpenTemplatePackage("CalculationTwr.xlsx"))
            {
                var ws = p.Workbook.Worksheets[6];
                var dp = OptimizedDependencyChainFactory.Create(ws.Cells, new ExcelCalculationOption() { });
                //Assert.AreEqual(1, dp._circularReferences.Count);
            }
        }

    }
}
