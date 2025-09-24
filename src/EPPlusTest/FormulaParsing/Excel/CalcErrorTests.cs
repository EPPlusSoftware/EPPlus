using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel
{
    [TestClass]
    public class CalcErrorTests : TestBase
    {
        [TestMethod]
        public void CalcErrorWhenEmptyFilterFunction()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "Jack";
            sheet.Cells["B1"].Value = "ECE";
            sheet.Cells["C1"].Value = 5;
            sheet.Cells["A2"].Value = "Adam";
            sheet.Cells["B2"].Value = "CSE";
            sheet.Cells["C2"].Value = 8;
            sheet.Cells["A3"].Value = "Julie";
            sheet.Cells["B3"].Value = "ECE";
            sheet.Cells["C3"].Value = 7;
            sheet.Cells["E4"].Formula = "FILTER(A1:C3,C1:C3<4)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Calc), sheet.Cells["E4"].Value);
            //SaveWorkbook("CalcError1.xlsx", package);
        }

        [TestMethod]
        public void CalcErrorWhenUnInvokedLambda()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(x,x/0.83)";
            var tokens = SourceCodeTokenizer.Default.Tokenize("LAMBDA(x,x/0.83)");
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Calc), sheet.Cells["A1"].Value);
            //SaveWorkbook("CalcError2.xlsx", package);
        }
    }
}
