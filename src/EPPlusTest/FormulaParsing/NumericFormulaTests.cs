using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class NumericFormulaTests
    {
        [TestMethod]
        public void ShouldHandleNumericFormulaLow()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Formula = "4.00801603206413E-06";
                sheet.Calculate();
                Assert.AreEqual(4.00801603206413E-06d, sheet.Cells[1, 1].Value);
            }
        }

        [TestMethod]
        public void ShouldHandleNegativeNumericFormulaLow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Formula = "-4.00801603206413E-06";
                sheet.Calculate();
                Assert.AreEqual(-4.00801603206413E-06d, sheet.Cells[1, 1].Value);
            }
        }

        public void ShouldHandleNumericFormulaHigh()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Formula = "4.00801603206413E+06";
                sheet.Calculate();
                Assert.AreEqual(4.00801603206413E+06d, sheet.Cells[1, 1].Value);
            }
        }

        [TestMethod]
        public void ShouldHandleNegativeNumericFormulaHigh()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Formula = "-4.00801603206413E+06";
                sheet.Calculate();
                Assert.AreEqual(-4.00801603206413E+06d, sheet.Cells[1, 1].Value);
            }
        }

        [TestMethod]
        public void ShouldHandleNumericFormulaWithOperator()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Formula = "4.00801603206413E-06 + 1";
                sheet.Calculate();
                var result = Math.Round((double)sheet.Cells[1, 1].Value, 9);
                Assert.AreEqual(1.000004008, result);
            }
        }
    }
}
