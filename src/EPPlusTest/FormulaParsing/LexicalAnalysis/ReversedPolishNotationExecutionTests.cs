using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class ReversedPolishNotationExecutionTests
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
        }

        private void SetUpWorksheet1()
        {
            var ws1 = _package.Workbook.Worksheets.Add("Sheet2");
            ws1.Cells["A1:A1000"].Value = 1;
            ws1.Cells["B1"].Value = 100;
            ws1.Cells["B2:B1000"].Value = "B1*(A2/A1)+1";

            ws1.Cells["C1"].Formula = "B1000/B1-1";
            ws1.Cells["D1"].Formula = "Sum(B1:B1000)/C1";
        }

        private void SetUpWorksheet2()
        {
        //    var ws1 = _package.Workbook.Worksheets.Add("Sheet1");
        //    ws1.Cells["A1"].Value = 1;
        //    ws1.Cells["B1"].Value = 2;
        //    ws1.Cells["C1"].Value = 3;

        //    ws1.Cells["A2"].Value = 10;
        //    ws1.Cells["B2"].Value = 20;
        //    ws1.Cells["C2"].Value = 30;

        //    _package.Workbook.Names.AddValue("WorkbookDefinedNameValue", 1);
        //    ws1.Names.AddValue("WorksheetDefinedNameValue", "Name Value");
        }

        [TestMethod]
        public void ExecuteWorksheet1()
        {
            var dp = new RpnOptimizedDependencyChain(_package.Workbook);
            dp.Execute();
        }
    }
}
