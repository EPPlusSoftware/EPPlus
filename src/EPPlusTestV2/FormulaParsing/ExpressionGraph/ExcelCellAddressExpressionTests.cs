using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    //TODO: Look at these test and add support for the new formula parser.
    [TestClass]
    public class ExcelCellAddressTests : TestBase
    {
        static ExcelPackage _package;
        static EpplusExcelDataProvider _excelDataProvider;
        //static ExpressionGraphBuilder _graphBuilder;
        static ExcelWorksheet _ws, _ws2;
        internal static ISourceCodeTokenizer _tokenizer = SourceCodeTokenizer.Default;
        //static ExpressionCompiler _compiler;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _package = OpenPackage("CellAddressExpression.xlsx", true);

            _ws = _package.Workbook.Worksheets.Add("Sheet1");
            _ws2 = _package.Workbook.Worksheets.Add("Sheet2");
            LoadTestdata(_ws, 100,1,1,false,false, new DateTime(2022,11,1));
            var tbl = _ws.Tables.Add(_ws.Cells["A1:E101"], "MyTable");
            _package.Workbook.Names.Add("WorkbookName1", _ws.Cells["L15"]);
            _package.Workbook.Names.Add("WorkbookNameSheet2", _ws.Cells["L15"]);
            _ws.Names.Add("SingleCellName", _ws.Cells["H3"]);
            _ws.Names.Add("RangeName", _ws.Cells["G5:H8"]);
            tbl.ShowTotal = true;
            var parsingContext = ParsingContext.Create(_package);
            _excelDataProvider = new EpplusExcelDataProvider(_package, parsingContext);
            //_compiler = new ExpressionCompiler(parsingContext);

            //parsingContext.Scopes.NewScope(new FormulaRangeAddress() { WorksheetIx=0, FromRow = 1, FromCol = 1, ToRow = 1,ToCol = 1 });
            parsingContext.CurrentCell = new FormulaCellAddress(0, 1, 1);
            parsingContext.ExcelDataProvider = _excelDataProvider;
            //_graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, parsingContext);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_package);
            _package.Dispose();
        }
        [TestMethod]
        public void VerifyCellAddressExpression_NonFixed()
        {
            //Setup
            _ws.Cells["J1"].Formula = @"SUM(A1:C5)";
            _ws.Cells["J1"].Calculate();
            //Assert
            Assert.AreEqual(179484D, _ws.Cells["J1"].Value);
        }
        [TestMethod]
        public void VerifyCellAddressExpression_MultiColon()
        {
            //Setup
            _ws.Cells["J2"].Formula = @"SUM(Sheet1!A1:C5:E2)";
            _ws.Cells["J2"].Calculate();

            //Assert
            Assert.AreEqual(179946D, _ws.Cells["J2"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndNameSingeCell()
        {
            //Setup
            _ws.Cells["J3"].Formula = @"Sum(Sheet1!SingleCellName:A1)";
            _ws.Cells["J3"].Calculate();
            //Assert
            Assert.AreEqual(89903D, _ws.Cells["J3"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndNameSingeCell_Reversed()
        {
            //Setup
            _ws.Cells["J4"].Formula = @"Sum(Sheet1!A1:SingleCellName)";
            _ws.Cells["J4"].Calculate();
            //Assert
            Assert.AreEqual(89903D, _ws.Cells["J4"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndNameRange()
        {
            //Setup
            _ws.Cells["J5"].Formula = @"Sum(Sheet1!RangeName:B6)";
            _ws.Cells["J5"].Calculate();

            //Assert
            Assert.AreEqual(884D, _ws.Cells["J5"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndNameRange_Reversed()
        {
            //Setup
            _ws.Cells["J6"].Formula = @"Sum(Sheet1!B6:RangeName)";
            _ws.Cells["J6"].Calculate();

            //Assert
            Assert.AreEqual(884D, _ws.Cells["J6"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndWorkbookNameRange()
        {
            //Setup
            _ws.Cells["J7"].Formula = @"Sum(Sheet1!D17:WorkbookName1)";
            _ws.Cells["J7"].Calculate();

            //Assert
            Assert.AreEqual(1584D, _ws.Cells["J7"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndWorkbookNameRange_Reverse()
        {
            //Setup
            _ws.Cells["J8"].Formula = @"Sum(WorkbookName1:Sheet1!D17)";
            _ws.Cells["J8"].Calculate();
            
            //Assert
            Assert.AreEqual(1584D, _ws.Cells["J8"].Value);
        }
        [TestMethod]
        public void VerifyRangeAndWorkbookFunction()
        {
            //Setup
            _ws.Cells["J9"].Formula = @"Sum(Offset(A1,1,1):Sheet1!I15)";
            _ws.Cells["J9"].Calculate();

            //Assert
            Assert.AreEqual(4046D, _ws.Cells["J9"].Value);
        }
        [TestMethod]
        public void VerifyRangeExpressionInFunction()
        {
            //Setup
            _ws.Cells["J10"].Formula = @"Sum(IF(FALSE,A1:A2,B1:B2))";
            _ws.Cells["J10"].Calculate();

            //Assert
            Assert.AreEqual(2D, _ws.Cells["J10"].Value);
        }
    }
}
