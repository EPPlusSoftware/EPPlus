using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class LambdaFunctionTests
    {
        [TestMethod]
        public void LambdaSelfInvokeTest1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(r,c,r+c)(D6,D7)";
            sheet.Cells["D6"].Value = 5;
            sheet.Cells["D7"].Value = 6;
            sheet.Calculate();
            Assert.AreEqual(11d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LambdaSelfInvokeTest2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "IF(TRUE(),LAMBDA(r,c,r-c),1)(D6,D7)";
            sheet.Cells["D6"].Value = 7;
            sheet.Cells["D7"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(5d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LambdaRecursive1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(a,a + LAMBDA(b,b + a)(2))(2)";
            sheet.Calculate();
            Assert.AreEqual(6d, sheet.Cells["A1"].Value);
        }
    }
}
