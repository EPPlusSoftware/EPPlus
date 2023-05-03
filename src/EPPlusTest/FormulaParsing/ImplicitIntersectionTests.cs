using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Packaging.Ionic.Crc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ImplicitIntersectionTests : TestBase
    {
        [TestMethod]
        public void SingleShouldReturnValueErrorWhenMultipleRowsAndColsInRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 1;
                sheet.Cells[2, 1].Value = 2;
                sheet.Cells[3, 1].Value = 3;
                sheet.Cells[1, 2].Value = 4;
                sheet.Cells[2, 2].Value = 5;
                sheet.Cells[3, 2].Value = 6;
                sheet.Cells["C3"].Formula = "_xlfn.SINGLE(A1:B3)";
                sheet.Calculate();
                //SaveAndCleanup(package);
                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["C3"].Value);
            }
        }

        [TestMethod]
        public void SingleShouldReturnValueErrorWhenRowIsOutOfRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[2, 2].Value = 1;
                sheet.Cells[3, 2].Value = 2;
                sheet.Cells[4, 2].Value = 3;
                sheet.Cells[2, 3].Value = 4;
                sheet.Cells[3, 3].Value = 5;
                sheet.Cells[4, 3].Value = 6;
                sheet.Cells["C6"].Formula = "_xlfn.SINGLE(B1:B3)";
                sheet.Cells["A4"].Formula = "_xlfn.SINGLE(B4:C4)"; ;
                sheet.Calculate();
                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["C6"].Value, "C6 was not #VALUE");
            }
        }

        [TestMethod]
        public void SingleShouldReturnValueErrorWhenMultipleOptionsInDirection()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[2, 2].Value = 1;
                sheet.Cells[3, 2].Value = 2;
                sheet.Cells[4, 2].Value = 3;
                sheet.Cells[2, 3].Value = 4;
                sheet.Cells[3, 3].Value = 5;
                sheet.Cells[4, 3].Value = 6;
                sheet.Cells["B1"].Formula = "_xlfn.SINGLE(B2:B4)";
                sheet.Cells["B5"].Formula = "_xlfn.SINGLE(B2:B4)";
                sheet.Cells["D2"].Formula = "_xlfn.SINGLE(B2:C2)";
                sheet.Calculate();
                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["B1"].Value, "B1 was not #VALUE");
                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["B5"].Value, "B5 was not #VALUE");
                Assert.AreEqual(ErrorValues.ValueError, sheet.Cells["D2"].Value, "D2 was not #VALUE");
            }
        }

        [TestMethod]
        public void SingleShouldReturnValueFromCorrectCell()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[2, 2].Value = 1;
                sheet.Cells[3, 2].Value = 2;
                sheet.Cells[4, 2].Value = 3;
                sheet.Cells[2, 3].Value = 4;
                sheet.Cells[3, 3].Value = 5;
                sheet.Cells[4, 3].Value = 6;
                // to the right
                sheet.Cells["D3"].Formula = "_xlfn.SINGLE(C2:C4)";
                // to the left
                sheet.Cells["A3"].Formula = "_xlfn.SINGLE(B2:B4)";
                // above
                sheet.Cells["C1"].Formula = "_xlfn.SINGLE(B2:C2)";
                // underneath
                sheet.Cells["C5"].Formula = "_xlfn.SINGLE(B4:C4)";
                sheet.Calculate();
                Assert.AreEqual(5, sheet.Cells["D3"].Value, "D3 was not 5");
                Assert.AreEqual(2, sheet.Cells["A3"].Value, "A3 was not 2");
                Assert.AreEqual(4, sheet.Cells["C1"].Value, "C1 was not 4");
                Assert.AreEqual(6, sheet.Cells["C5"].Value, "C5 was not 6");
            }
        }

        [TestMethod]
        public void SingleShouldReturnValueFromCorrectCell_FromOtherWs()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[2, 2].Value = 1;
                sheet.Cells[3, 2].Value = 2;
                sheet.Cells[4, 2].Value = 3;
                sheet.Cells[2, 3].Value = 4;
                sheet.Cells[3, 3].Value = 5;
                sheet.Cells[4, 3].Value = 6;
                var sheet2 = package.Workbook.Worksheets.Add("test2");
                // to the right
                sheet2.Cells["D3"].Formula = "_xlfn.SINGLE(test!C2:C4)";
                // to the left
                sheet2.Cells["A3"].Formula = "_xlfn.SINGLE(test!B2:B4)";
                // above
                sheet2.Cells["C1"].Formula = "_xlfn.SINGLE(test!B2:C2)";
                // underneath
                sheet2.Cells["C5"].Formula = "_xlfn.SINGLE(test!B4:C4)";

                sheet2.Calculate();
                Assert.AreEqual(5, sheet2.Cells["D3"].Value, "D3 was not 5");
                Assert.AreEqual(2, sheet2.Cells["A3"].Value, "A3 was not 2");
                Assert.AreEqual(4, sheet2.Cells["C1"].Value, "C1 was not 4");
                Assert.AreEqual(6, sheet2.Cells["C5"].Value, "C5 was not 6");
            }
        }

        [TestMethod]
        public void ShouldUseIntersectionCellInAddress()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Formula = "IF(1>0,_xlfn.SINGLE(A5:B5),\"A1\"):B6";
                sheet.Cells[5, 1].Value = 1;
                sheet.Cells[6, 1].Value = 2;
                sheet.Cells[5, 2].Value = 3;
                sheet.Cells[6, 2].Value = 4;

                sheet.Calculate();

                Assert.AreEqual(1, sheet.Cells["A1"].Value);
                Assert.AreEqual(2, sheet.Cells["A2"].Value);
                Assert.AreEqual(3, sheet.Cells["B1"].Value);
                Assert.AreEqual(4, sheet.Cells["B2"].Value);
            }
        }

        [TestMethod, Ignore]
        public void SingleShouldDoImplicitIntersection()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 2].Value = 2;
                sheet.Cells[3, 2].Value = 3;
                sheet.Cells[4, 2].Value = 4;
                sheet.Cells[5, 2].Value = 5;
                sheet.Cells["A1:A5"].Formula = "B1:B5";
                sheet.Cells["C3"].Formula = "_xlfn.SINGLE(B1:B3)";
                sheet.Calculate();
                Assert.AreEqual(3, sheet.Cells["C3"].Value);
            }
        }
    }
}
