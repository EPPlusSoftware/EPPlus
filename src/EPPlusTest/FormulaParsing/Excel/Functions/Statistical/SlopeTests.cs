using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class SlopeTests
    {
        [TestMethod]
        public void SlopeTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["D2"].Value = new DateTime(1900, 1, 2);
                sheet.Cells["E2"].Value = 6;
                sheet.Cells["D3"].Value = new DateTime(1900, 1, 3);
                sheet.Cells["E3"].Value = 5;
                sheet.Cells["D4"].Value = new DateTime(1900, 1, 9);
                sheet.Cells["E4"].Value = 11;
                sheet.Cells["D5"].Value = new DateTime(1900, 1, 1);
                sheet.Cells["E5"].Value = 7;
                sheet.Cells["D6"].Value = new DateTime(1900, 1, 8);
                sheet.Cells["E6"].Value = 5;
                sheet.Cells["D7"].Value = new DateTime(1900, 1, 7);
                sheet.Cells["E7"].Value = 4;
                sheet.Cells["D8"].Value = new DateTime(1900, 1, 5);
                sheet.Cells["E8"].Value = 4;

                sheet.Cells["E9"].Formula = "SLOPE(D2:D8,E2:E8)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E9"].Value, 7);
                Assert.AreEqual(0.3055556, result);
            }
        }

        [TestMethod]
        public void SlopeShouldCalculateWithEmptyCell()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["D2"].Value = new DateTime(1900, 1, 2);
                sheet.Cells["E2"].Value = 6;
                sheet.Cells["D3"].Value = new DateTime(1900, 1, 3);
                sheet.Cells["E3"].Value = 5;
                sheet.Cells["D4"].Value = new DateTime(1900, 1, 9);
                sheet.Cells["E4"].Value = 11;
                sheet.Cells["D5"].Value = new DateTime(1900, 1, 1);
                sheet.Cells["E5"].Value = 7;
                sheet.Cells["D6"].Value = new DateTime(1900, 1, 8);
                sheet.Cells["E6"].Value = 5;
                sheet.Cells["D7"].Value = new DateTime(1900, 1, 7);
                sheet.Cells["E7"].Value = 4;
                sheet.Cells["D8"].Value = new DateTime(1900, 1, 5);
                //Cell E8 is empty
                sheet.Cells["E9"].Formula = "SLOPE(D2:D8,E2:E8)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E9"].Value, 7);
                Assert.AreEqual(0.3510638, result);
            }
        }

        [TestMethod, Ignore]
        public void SlopeShouldCalculateDiffrentRangeForms()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["D2"].Value = 2;
                sheet.Cells["H2"].Value = 1;
                sheet.Cells["E2"].Value = 5;
                sheet.Cells["I2"].Value = 2;
                sheet.Cells["D3"].Value = 3;
                sheet.Cells["J2"].Value = 3;
                sheet.Cells["E3"].Value = 4;
                sheet.Cells["K2"].Value = 4;
                sheet.Cells["D4"].Value = 4;
                sheet.Cells["L2"].Value = 5;
                sheet.Cells["E4"].Value = 7;
                sheet.Cells["M2"].Value = 6;
 
                sheet.Cells["E9"].Formula = "SLOPE(D2:E4, H2:M2)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E9"].Value, 7);
                Assert.AreEqual(1.348720734, result);
            }
        }

        [TestMethod]
        public void SlopeShouldCalculate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["B2"].Value = 5;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["B4"].Value = 7;

                sheet.Cells["D5"].Value = 1;
                sheet.Cells["E5"].Value = 2;
                sheet.Cells["F5"].Value = 3;
                sheet.Cells["G5"].Value = 4;
                sheet.Cells["H5"].Value = 5;
                sheet.Cells["I5"].Value = 6;

                sheet.Cells["E9"].Formula = "SLOPE(A2:B4,D5:I5)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["E9"].Value, 9);
                Assert.AreEqual(0.657142857, result);
            }
        }
    }
}
