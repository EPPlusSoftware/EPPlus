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
            _package = new ExcelPackage();

            _ws = _package.Workbook.Worksheets.Add("Sheet1");
            _ws2 = _package.Workbook.Worksheets.Add("Sheet2");
            LoadTestdata(_ws);
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
        public void Cleanup()
        {
            SaveWorkbook("CellAddressExpression.xlsx", _package);
            _package.Dispose();
        }
        [TestMethod]
        public void VerifyCellAddressExpression_NonFixed()
        {
            //Setup
            var f = @"SUM(A1:C5)";
            var tokens = _tokenizer.Tokenize(f);
            //RpnFormulaExecution.Execute()
            //var exps = _graphBuilder.Build(tokens);

            //Assert
            var r = RpnFormulaExecution.ExecuteFormula(_package.Workbook, f, new ExcelCalculationOption());
            Assert.AreEqual(6, tokens.Count);
        //    Assert.AreEqual(1, exps.Expressions.Count);

        //    Assert.AreEqual(TokenType.CellAddress, tokens[2].TokenType);
        //    Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
        //    var ra = exps.Expressions[0].Children[0].Children[0];
        //    Assert.AreEqual(2, ra.Children.Count);
        //    var result1 = ((CellAddressExpression)ra.Children[0]).Compile();
        //    var result2 = ((CellAddressExpression)ra.Children[1]).Compile();
        //    var range1 = (FormulaRangeAddress)result1.Result;
        //    var range2 = (FormulaRangeAddress)result2.Result;

        //    Assert.AreEqual(range1.FromRow, 1);
        //    Assert.AreEqual(range1.FromCol, 1);
        //    //Assert.AreEqual(range1.FixedFlag, FixedFlag.None);
        //    Assert.AreEqual(range2.FromRow, 5);
        //    Assert.AreEqual(range2.FromCol, 3);
        //    //Assert.AreEqual(range2.FixedFlag, FixedFlag.None);
        }
        [TestMethod]
        public void VerifyCellAddressExpression_MultiColon()
        {
            //Setup
            var f = @"Sheet1!A1:C5:E2";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);
            //_package.Workbook.Worksheets[0].Cells["H5"].Formula = f;

            //Assert
            Assert.AreEqual(7, tokens.Count);
            //Assert.AreEqual(3, expTree.Expressions[0].Children.Count);

            //Assert.AreEqual(TokenType.CellAddress, tokens[2].TokenType);
            //Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
            //Assert.AreEqual(TokenType.CellAddress, tokens[6].TokenType);
            //var range = (IRangeInfo)result.Result;

            //Assert.AreEqual(range.Address.ExternalReferenceIx, -1);
            //Assert.AreEqual(range.Address.WorksheetIx, 0);

            //Assert.AreEqual(range.Address.FromRow, 1);
            //Assert.AreEqual(range.Address.FromCol, 1);
            //Assert.AreEqual(range.Address.ToRow, 5);
            //Assert.AreEqual(range.Address.ToCol, 5);
            //Assert.AreEqual(range.Address.FixedFlag, FixedFlag.None);
        }
        [TestMethod]
        public void VerifyCellAddressExpression_Fixed()
        {
            //Setup
            var f = @"[0]Sheet1!A1:C5:E2";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);
            //_package.Workbook.Worksheets[0].Cells["H5"].Formula = f;

            //Assert
            Assert.AreEqual(10, tokens.Count);
            //Assert.AreEqual(3, expTree.Expressions[0].Children.Count);

            //Assert.AreEqual(TokenType.CellAddress, tokens[5].TokenType);
            //Assert.AreEqual(TokenType.CellAddress, tokens[7].TokenType);
            //Assert.AreEqual(TokenType.CellAddress, tokens[9].TokenType);
            //var range = (IRangeInfo)result.Result;

            //Assert.AreEqual(range.Address.ExternalReferenceIx, -1);
            //Assert.AreEqual(range.Address.WorksheetIx, 0);

            //Assert.AreEqual(range.Address.FromRow, 1);
            //Assert.AreEqual(range.Address.FromCol, 1);
            //Assert.AreEqual(range.Address.ToRow, 5);
            //Assert.AreEqual(range.Address.ToCol, 5);
            //Assert.AreEqual(range.Address.FixedFlag, FixedFlag.None);
        }
        [TestMethod]
        public void VerifyCellAddressExpression_WithSum()
        {
            //Setup
            var f = @"Sum(Sheet1!A1:C5:E2:A39 + Sheet1!F1:J39)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);
            _package.Workbook.Worksheets[0].Cells["H5"].Formula = f;

            //Assert
            Assert.AreEqual(18, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children.Count);

            Assert.AreEqual(TokenType.CellAddress, tokens[4].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[6].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[8].TokenType);
            Assert.AreEqual(TokenType.CellAddress, tokens[10].TokenType);
            //var range1 = (FormulaRangeAddress)expTree.Expressions[0].Children[0].Children[0].Compile().Result;
            //var range2 = (FormulaRangeAddress)expTree.Expressions[0].Children[0].Children[1].Compile().Result;

            ////Assert Range 1
            //Assert.AreEqual(range1.ExternalReferenceIx, -1);
            //Assert.AreEqual(range1.WorksheetIx, 0);
            //Assert.AreEqual(range1.FromRow, 1);
            //Assert.AreEqual(range1.FromCol, 1);
            //Assert.AreEqual(range1.ToRow, 39);
            //Assert.AreEqual(range1.ToCol, 5);
            //Assert.AreEqual(range1.FixedFlag, FixedFlag.None);

            //Assert Range 2
            //Assert.AreEqual(range2.ExternalReferenceIx, -1);
            //Assert.AreEqual(range2.WorksheetIx, 0);
            //Assert.AreEqual(range2.FromRow, 1);
            //    Assert.AreEqual(range2.FromCol, 6);
            //    Assert.AreEqual(range2.ToRow, 39);
            //Assert.AreEqual(range2.ToCol, 10);
                                 
            //Assert.AreEqual(range2.FixedFlag, FixedFlag.None);
        }
        [TestMethod]
        public void VerifyRangeAndNameSingeCell()
        {
            //Setup
            var f = @"Sum(Sheet1!SingleCellName:A1)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(1, range.FromRow);
            //Assert.AreEqual(1, range.FromCol);
            //Assert.AreEqual(3, range.ToRow);
            //Assert.AreEqual(8, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeAndNameSingeCell_Reversed()
        {
            //Setup
            var f = @"Sum(Sheet1!A1:SingleCellName)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(1, range.FromRow);
            //Assert.AreEqual(1, range.FromCol);
            //Assert.AreEqual(3, range.ToRow);
            //Assert.AreEqual(8, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeAndNameRange()
        {
            //Setup
            var f = @"Sum(Sheet1!RangeName:B6)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(5, range.FromRow);
            //Assert.AreEqual(2, range.FromCol);
            //Assert.AreEqual(8, range.ToRow);
            //Assert.AreEqual(8, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeAndNameRange_Reversed()
        {
            //Setup
            var f = @"Sum(Sheet1!B6:RangeName)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(5, range.FromRow);
            //Assert.AreEqual(2, range.FromCol);
            //Assert.AreEqual(8, range.ToRow);
            //Assert.AreEqual(8, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeAndWorkbookNameRange()
        {
            //Setup
            var f = @"Sum(Sheet1!J15:WorkbookName1)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            ////Assert
            //Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(15, range.FromRow);
            //Assert.AreEqual(10, range.FromCol);
            //Assert.AreEqual(15, range.ToRow);
            //Assert.AreEqual(12, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeAndWorkbookNameRange_Reverse()
        {
            //Setup
            var f = @"Sum(WorkbookName1:Sheet1!J15)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(15, range.FromRow);
            //Assert.AreEqual(10, range.FromCol);
            //Assert.AreEqual(15, range.ToRow);
            //Assert.AreEqual(12, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeAndWorkbookFunction()
        {
            //Setup
            var f = @"Sum(Offset(A1,1,1):Sheet1!J15)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(15, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(2, expTree.Expressions[0].Children[0].Children[0].Children.Count);

            //var resultRange = expTree.Expressions[0].Children[0].Children[0].Compile();
            //var range = (FormulaRangeAddress)resultRange.Result;
            //Assert.AreEqual(2, range.FromRow);
            //Assert.AreEqual(2, range.FromCol);
            //Assert.AreEqual(15, range.ToRow);
            //Assert.AreEqual(10, range.ToCol);
        }
        [TestMethod]
        public void VerifyRangeExpressionInFunction()
        {
            //Setup
            var f = @"IF(FALSE,A1:A2,B1:B2)";
            var tokens = _tokenizer.Tokenize(f);
            //var expTree = _graphBuilder.Build(tokens);
            //var result = _compiler.Compile(expTree.Expressions);

            //Assert
            Assert.AreEqual(12, tokens.Count);
            //Assert.AreEqual(1, expTree.Expressions.Count);
            //Assert.AreEqual(3, expTree.Expressions[0].Children.Count);
            //Assert.AreEqual(ExpressionType.Boolean, expTree.Expressions[0].Children[0].ExpressionType);
            //Assert.AreEqual(ExpressionType.RangeAddress, expTree.Expressions[0].Children[1].ExpressionType);
            //Assert.AreEqual(ExpressionType.RangeAddress, expTree.Expressions[0].Children[2].ExpressionType);
        }
    }
}
