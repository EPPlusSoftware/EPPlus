using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
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
        public void CircularRef_In_Sum_BetweenTwoCells_ShouldThow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Formula = "B2";
                sheet.Cells["B2"].Formula = "A2";
                sheet.Cells["A3"].Formula = "SUM(A1:A2)";
                sheet.Calculate();
            }
        }

        [TestMethod]
        public void CircularRef_In_Sum_BetweenTwoCells_ShouldThow_WhenAllow()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Formula = "B2";
                sheet.Cells["B2"].Formula = "A2";
                sheet.Cells["A3"].Formula = "SUM(A1:A2)";
                var calcOptions = new ExcelCalculationOption { AllowCircularReferences = true };
                sheet.Calculate(calcOptions);
                Assert.AreEqual(1d, sheet.Cells["A3"].Value);
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

        [TestMethod]
        public void VLookupShouldNotThrowWhenCircularRefsAllowed()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Formula = "B2";
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["B4"].Formula = "VLOOKUP(3, A1:B3, 2)";
                var calcOptions = new ExcelCalculationOption { AllowCircularReferences = true };
                sheet.Calculate(calcOptions);
                Assert.AreEqual(6, sheet.Cells["B4"].Value);
            }
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void VLookupShouldThrowWhenCircularRefsNotAllowed()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B1"].Value = 4;
                sheet.Cells["B2"].Formula = "B2";
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["B4"].Formula = "VLOOKUP(3, A1:B3, 2)";
                sheet.Calculate();
            }
        }

        [TestMethod]
        public void IfShouldIgnoreCircularRefWhenIgnoredArg()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2;
                sheet.Cells["B2"].Formula = "B2";
                sheet.Cells["B3"].Value = 6;
                sheet.Cells["B4"].Formula = "IF(A1<>2, B2, B3)";
                var calcOptions = new ExcelCalculationOption { AllowCircularReferences = true };
                sheet.Calculate(calcOptions);
                Assert.AreEqual(6, sheet.Cells["B4"].Value);
            }
        }
        [TestMethod]
        public void ReferenceToOtherSheetShouldNotGenerateACircularReference()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");

                sheet1.Cells["A1"].Value = 1d;
                sheet1.Cells["A2"].Value = 3d;
                sheet1.Cells["A3"].Formula = "=Sheet2!A1";

                sheet2.Cells["A1"].Formula = "=Sheet1!A1";

                package.Workbook.Calculate();
            }
        }

    }
}
