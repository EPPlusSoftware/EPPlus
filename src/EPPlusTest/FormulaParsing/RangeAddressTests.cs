using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class RangeAddressTests
    {
        private ParsingContext _context;
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;
        private ExcelDataProvider _excelDataProvider;
        private ExpressionGraphBuilder _graphBuilder;
        private ExpressionCompiler _expressionCompiler;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
            _context = ParsingContext.Create(_package);
            _excelDataProvider = new EpplusExcelDataProvider(_package, _context);
            _context.ExcelDataProvider = _excelDataProvider;
            _graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, _context);
            _expressionCompiler = new ExpressionCompiler(_context);
            var currentAddress = new FormulaRangeAddress()
            {
                WorksheetIx = 0,
                FromCol = 0,
                FromRow = 10,
                ToCol = 0,
                ToRow = 10
            };
            _context.Scopes.NewScope(currentAddress);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _graphBuilder = null;
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldSetSimpleRangeAddress()
        {
            var input = "A1:A2";
            var tokens = OptimizedSourceCodeTokenizer.Default.Tokenize(input);
            var graph = _graphBuilder.Build(tokens);
            var compileResult = _expressionCompiler.Compile(graph.Expressions);
            Assert.AreEqual(DataType.ExcelRange, compileResult.DataType);
            Assert.IsInstanceOfType(compileResult.Result, typeof(IRangeInfo));
            
            var address = ((IRangeInfo)compileResult.Result).Address;
            Assert.AreEqual(0, address.WorksheetIx);
            Assert.AreEqual(1, address.FromRow);
            Assert.AreEqual(2, address.ToRow);
            Assert.AreEqual(1, address.FromCol);
            Assert.AreEqual(1, address.ToCol);
        }

        [TestMethod]
        public void ShouldSetRangeAddressOnOtherWorksheet()
        {
            _package.Workbook.Worksheets.Add("Sheet2");
            var input = "Sheet2!A1:A2";
            var tokens = OptimizedSourceCodeTokenizer.Default.Tokenize(input);
            var graph = _graphBuilder.Build(tokens);
            var compileResult = _expressionCompiler.Compile(graph.Expressions);
            Assert.AreEqual(DataType.ExcelRange, compileResult.DataType);
            Assert.IsInstanceOfType(compileResult.Result, typeof(IRangeInfo));

            var address = ((IRangeInfo)compileResult.Result).Address;
            Assert.AreEqual(1, address.WorksheetIx);
            Assert.AreEqual(1, address.FromRow);
            Assert.AreEqual(2, address.ToRow);
            Assert.AreEqual(1, address.FromCol);
            Assert.AreEqual(1, address.ToCol);
        }

        [TestMethod]
        public void ShouldSetSingleCellAddress()
        {
            var input = "A1";
            var tokens = OptimizedSourceCodeTokenizer.Default.Tokenize(input);
            var graph = _graphBuilder.Build(tokens);
            var compileResult = _expressionCompiler.Compile(graph.Expressions);
            Assert.AreEqual(DataType.Empty, compileResult.DataType);
            Assert.IsNull(compileResult.Result);
            var address = compileResult.Address;
            Assert.AreEqual(0, address.WorksheetIx);
            Assert.AreEqual(1, address.FromCol);
            Assert.AreEqual(1, address.FromRow);
            Assert.AreEqual(1, address.ToCol);
            Assert.AreEqual(1, address.ToRow);
        }

        [TestMethod]
        public void ShouldSetSingleCellAddressOnOtherWorksheet()
        {
            _package.Workbook.Worksheets.Add("Sheet2");
            var input = "Sheet2!A1";
            var tokens = OptimizedSourceCodeTokenizer.Default.Tokenize(input);
            var graph = _graphBuilder.Build(tokens);
            var compileResult = _expressionCompiler.Compile(graph.Expressions);
            Assert.AreEqual(DataType.Empty, compileResult.DataType);
            Assert.IsNull(compileResult.Result);
            var address = compileResult.Address;
            Assert.AreEqual(1, address.WorksheetIx);
            Assert.AreEqual(1, address.FromCol);
            Assert.AreEqual(1, address.FromRow);
            Assert.AreEqual(1, address.ToCol);
            Assert.AreEqual(1, address.ToRow);
        }
    }
}
