using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExcelTableExpressionTests : TestBase
    {
        static ExcelPackage _package;
        static ExcelWorksheet _ws;
        internal static ISourceCodeTokenizer _tokenizer = SourceCodeTokenizer.Default;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _package = OpenPackage("TableExpression.xlsx", true);

            _ws = _package.Workbook.Worksheets.Add("Sheet1");
            LoadTestdata(_ws);
            var tbl = _ws.Tables.Add(_ws.Cells["A1:E101"], "MyTable");
            tbl.ShowTotal = true;
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_package);
            _package.Dispose();
        }
        [TestMethod]
        public void VerifyTableExpression_Table()
        {
            //Setup
            var f = @"SUM(MyTable[])";

            _ws.Cells["H1"].Formula = f;
            _ws.Cells["H1"].Calculate();

            Assert.AreEqual(4618251D, _ws.Cells["H1"].Value);
        }
        [TestMethod]
        public void VerifyTableExpression_All_Column_Date()
        {
            //Setup
            var f = @"SUM(MyTable[[#all],[Date]])";
            _ws.Cells["H2"].Formula = f;
            _ws.Cells["H2"].Calculate();

            Assert.AreEqual(4446585D, _ws.Cells["H2"].Value);
        }
        [TestMethod]
        public void VerifyTableExpression_Header_Data_Column_Date()
        {
            //Setup
            var f = @"SUM(MyTable[[#Headers],[#Data],[Date]])";
            _ws.Cells["H3"].Formula = f;
            _ws.Cells["H3"].Calculate();

            Assert.AreEqual(4446585D, _ws.Cells["H3"].Value);
        }
        [TestMethod]
        public void VerifyTableExpression_Data_Totals_Column_Date()
        {
            //Setup
            var f = @"SUM(MyTable[[#DATA],[#Totals],[NumFormattedValuE]])";
            _ws.Cells["H4"].Formula = f;
            _ws.Cells["H4"].Calculate();

            Assert.AreEqual(166617D, _ws.Cells["H4"].Value);
        }
        [TestMethod]
        public void VerifyTableExpression_all_Data_to_StrValue()
        {
            //Setup
            var f = @"SUM(MyTable[[#all],[Date]:[StrValue]])";
            _ws.Cells["H5"].Formula = f;
            _ws.Cells["H5"].Calculate();

            Assert.AreEqual(4451634D, _ws.Cells["H5"].Value);
        }
        [TestMethod]
        public void VerifyTableExpression_Table_With_Worksheet()
        {
            //Setup
            var f = @"SUM(Sheet1!MyTable[])";

            //_ws.Cells["G1"].Formula = f;
            //var tokens = _tokenizer.Tokenize(f);
            //var exps = _graphBuilder.Build(tokens);

            ////Assert
            //Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, exps.Expressions.Count);

            //Assert.AreEqual(TokenType.TableName, tokens[4].TokenType);
            //var result = ((TableAddressExpression)exps.Expressions[0].Children[0].Children[0]).Compile();
            //var range = (IRangeInfo)result.Result;

            //Assert.AreEqual(range.Address.FromRow, 2);
            //Assert.AreEqual(range.Address.FromCol, 1);
            //Assert.AreEqual(range.Address.ToRow, 101);
            //Assert.AreEqual(range.Address.ToCol, 5);
            //Assert.AreEqual(range.Address.FixedFlag, FixedFlag.All);
        }
        [TestMethod]


        public void VerifyTableExpression_Table_With_NonExisting_Worksheet()
        {
            //Setup
            var f = @"SUM(Sheet2!MyTable[])";

            _ws.Cells["H6"].Formula = f;
            _ws.Cells["H6"].Calculate();

            Assert.IsInstanceOfType(_ws.Cells["H6"].Value, typeof(ExcelErrorValue));

            //_ws.Cells["G1"].Formula = f;
            //var tokens = _tokenizer.Tokenize(f);
            //var exps = _graphBuilder.Build(tokens);

            ////Assert
            //Assert.AreEqual(8, tokens.Count);
            //Assert.AreEqual(1, exps.Expressions.Count);

            //Assert.AreEqual(TokenType.TableName, tokens[4].TokenType);
            //var result = ((TableAddressExpression)exps.Expressions[0].Children[0].Children[0]).Compile();
            //var range = (IRangeInfo)result.Result;

            //Assert.AreEqual("#REF!", range.Address.WorksheetAddress);
        }
        [TestMethod]
        public void VerifyTableExpression_External_Table()
        {
            //Setup
            var f = @"SUM([0]Sheet1!MyTable[])";
            _ws.Cells["H7"].Formula = f;
            _ws.Cells["H7"].Calculate();

            Assert.AreEqual(4618251D, _ws.Cells["H7"].Value);

            //_ws.Cells["G1"].Formula = f;
            //var tokens = _tokenizer.Tokenize(f);
            //var exps = _graphBuilder.Build(tokens);

            ////Assert
            //Assert.AreEqual(11, tokens.Count);
            //Assert.AreEqual(1, exps.Expressions.Count);

            //Assert.AreEqual(TokenType.TableName, tokens[7].TokenType);
            //var result = ((TableAddressExpression)exps.Expressions[0].Children[0].Children[0]).Compile();
            //var range = (IRangeInfo)result.Result;

            //Assert.AreEqual(range.Address.FromRow, 2);
            //Assert.AreEqual(range.Address.FromCol, 1);
            //Assert.AreEqual(range.Address.ToRow, 101);
            //Assert.AreEqual(range.Address.ToCol, 5);
            //Assert.AreEqual(range.Address.FixedFlag, FixedFlag.All);
        }
        [TestMethod]
        public void VerifyTableExpression_Table_And_CellAddress()
        {
            //Setup
            var f = @"SUM(Sheet1!MyTable[]:F5)";

            _ws.Cells["H8"].Formula = f;
            _ws.Cells["H8"].Calculate();

            Assert.AreEqual(4618251D, _ws.Cells["H8"].Value);

            //_ws.Cells["G1"].Formula = f;
            //var tokens = _tokenizer.Tokenize(f);
            //var exps = _graphBuilder.Build(tokens);
            //var restult = _compiler.Compile(exps.Expressions);

            ////Assert
            //Assert.AreEqual(10, tokens.Count);
            //Assert.AreEqual(1, exps.Expressions.Count);

            //Assert.AreEqual(TokenType.TableName, tokens[4].TokenType);
            //var addressResult = exps.Expressions[0].Children[0].Children[0].Compile();

            //Assert.IsInstanceOfType(addressResult.Result, typeof(FormulaRangeAddress));
            //var rangeResult = (FormulaRangeAddress)addressResult.Result;
            //Assert.AreEqual(0, rangeResult.WorksheetIx);
            //Assert.AreEqual(1, rangeResult.FromCol);
            //Assert.AreEqual(2, rangeResult.FromRow);
            //Assert.AreEqual(101, rangeResult.ToRow);
            //Assert.AreEqual(7, rangeResult.ToCol);
        }
    }
}
