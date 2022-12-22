using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class ReversedPolishNotationTests
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

        private void SetUpWorksheet2()
        {
            var ws2 = _package.Workbook.Worksheets.Add("Sheet2");
            ws2.Cells["A1:A3"].Value = 4;

            ws2.Cells["C2"].Value = "Col 1";
            ws2.Cells["D2"].Value = "Col 2";

            ws2.Cells["C3"].Value = 1;
            ws2.Cells["D3"].Value = new DateTime(2022, 10, 01);

            ws2.Cells["C4"].Value = 2;
            ws2.Cells["D4"].Value = new DateTime(2022, 11, 01);


            ws2.Tables.Add(ws2.Cells["C2:D4"], "Table1");
        }

        private void SetUpWorksheet1()
        {
            var ws1 = _package.Workbook.Worksheets.Add("Sheet1");
            ws1.Cells["A1"].Value = 1;
            ws1.Cells["B1"].Value = 2;
            ws1.Cells["C1"].Value = 3;

            ws1.Cells["A2"].Value = 10;
            ws1.Cells["B2"].Value = 20;
            ws1.Cells["C2"].Value = 30;

            _package.Workbook.Names.AddValue("WorkbookDefinedNameValue", 1);
            ws1.Names.AddValue("WorksheetDefinedNameValue", "Name Value");
        }

        [TestMethod]
        public void Calculate_NumericExpression1()
        {
            var formula = "3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());
            var expected = 3.001953125D;
            Assert.AreEqual(expected, value);
        }
        [TestMethod]
        public void Calculate_NumericExpression2()
        {
            var formula = "(( 1 -(- 2)-( 3 + 4 + 5 ))/( 6 + 7 * 8 - 9) * 10 )";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());
            var expected = -1.6981132075471697D;
            Assert.AreEqual(expected, value);
        }
        [TestMethod]
        public void Calculate_NumericExpression3()
        {
            var formula = "( 1 + 2 ) * ( 3 / 4 ) ^ ( 5 + 6 )";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());
            Assert.AreEqual(0.12670540809631348D, value);
        }
        [TestMethod]
        public void Calculate_NumericExpressionWithFunctions()
        {
            var formula = "sin(max((( 2 + 2 ) / 2), (3 * 3) / 3) / 3 * pi())";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());
            var expected = 3.231085104332676E-15;
            Assert.AreEqual(Math.Round(expected * 100000, 15), Math.Round((double)value * 100000, 15));
        }
        [TestMethod]
        public void Calculate_NumericExpressionWithAddresses1()
        {
            var formula = "A1 + B1 * C1 / ( 1 - 5 ) ^ 2 ^ 3";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            var expected = 1.00146484375;
            Assert.AreEqual(expected, value);
        }
        [TestMethod]
        public void Calculate_NumericExpressionWithAddresses2()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "(SUM(Sheet1!A1:C1)+1) * 3";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());

            Assert.AreEqual(21D, value);
        }
        [TestMethod]
        public void Calculate_NumericExpressionMultiplyTwoRanges()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "SUM(A1:B1+A2:B2)+1";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());

            Assert.AreEqual(34D, value);
        }
        [TestMethod]
        public void Calculate_Concat_Strings()
        {
            var formula = "\"Test\" & \" \" & \"2\"";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());
            Assert.AreEqual("Test 2", value);
        }
        [TestMethod]
        public void Calculate_Array()
        {
            var formula = "Sum({1,2;3,4})";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook, formula, new ExcelCalculationOption());
            Assert.AreEqual(10D, value);
        }
        [TestMethod]
        public void Calculate_ArrayAdditionWithRange()
        {
            var formula = "sum({1,2,3;3,4,5}+A1:C2)";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(84D, value);
        }
        [TestMethod]
        public void Calculate_TableColumnAddress()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "Sum(Table1[col 1])";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(3D, value);
        }
        [TestMethod]
        public void Calculate_AddressExpression_PreCompile()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 2, 4);
            var formula = "B1*(A2/A1)+1";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(21D, value);
        }

        [TestMethod]
        public void Calculate_Worksheet_NameFixedValue()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "Sheet1!WorksheetDefinedNameValue";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual("Name Value", value);
        }
        [TestMethod]
        public void Calculate_Workbook_NameFixedValue()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "WorkbookDefinedNameValue";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(1, value);
        }
        [TestMethod]
        public void Calculate_NonExisting_Worksheet_NameFixedValue()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "NonExistingSheet!WorksheetDefinedNameValue";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.IsInstanceOfType(value, typeof(ExcelErrorValue));
            Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)value).Type);
        }
        [TestMethod]
        public void FunctionsInFunctionsTests()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "IF((A1 * 1) > (A2 / 2) + 1,SUM(A1:C1),SUM(A2:C2))";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(60D, value);
        }
        [TestMethod]
        public void SumFullColumn()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "Sum(A:A)";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(11D, value);
        }
        [TestMethod]
        public void SumFullRow()
        {
            _parsingContext.CurrentCell = new FormulaCellAddress(0, 4, 1);
            var formula = "Sum(2:2)";
            var value = RpnFormulaExecution.ExecuteFormula(_package.Workbook.Worksheets[0], formula, new ExcelCalculationOption());
            Assert.AreEqual(60D, value);
        }

    }
}
