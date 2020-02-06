using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class CircularReferenceTests
    {
        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void CircularRef_In_Sum_ShouldThow()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "SUM(A1)";
                sheet.Calculate();
            }
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void CircularRef_In_Address_ShouldThow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "A1";
                sheet.Calculate();
            }
        }

        [TestMethod]
        public void CircularRef_In_Row_ShouldNotThow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "ROW(A1)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void CircularRef_In_Rows_ShouldNotThow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "ROWS(A1)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void CircularRef_In_Column_ShouldNotThow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COLUMN(A1)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void CircularRef_In_Columns_ShouldNotThow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COLUMNS(A1)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }
    }
}
