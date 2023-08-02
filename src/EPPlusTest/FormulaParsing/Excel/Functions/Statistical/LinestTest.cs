using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class LinestTest
    {
        [TestMethod]
        public void LinestTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5, B2:B5,,FALSE)";
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["A8"].Value);
                Assert.AreEqual(1d, sheet.Cells["B8"].Value);
            }
        }

        [TestMethod]
        public void LinestTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 4;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5,B2:B5,FALSE)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A8"].Value, 9);
                Assert.AreEqual(2.310344828d, result);
                Assert.AreEqual(0d, sheet.Cells["B8"].Value);
            }
        }

        [TestMethod]
        public void LinestTestWithStats()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test with statistics");
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 9;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 7;
                sheet.Cells["B2"].Value = 0;
                sheet.Cells["B3"].Value = 1;
                sheet.Cells["B4"].Value = 2;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A8"].Formula = "LINEST(A2:A5,B2:B5,TRUE,TRUE)";
                sheet.Calculate();
                var result1 = System.Math.Round((double)sheet.Cells["A8"].Value, 1);
                var result2 = System.Math.Round((double)sheet.Cells["B8"].Value, 1);
                var result3 = System.Math.Round((double)sheet.Cells["A9"].Value, 9);
                var result4 = System.Math.Round((double)sheet.Cells["B9"].Value, 9);
                var result5 = System.Math.Round((double)sheet.Cells["A10"].Value, 2);
                var result6 = System.Math.Round((double)sheet.Cells["B10"].Value, 9);
                var result7 = System.Math.Round((double)sheet.Cells["A11"].Value, 9);
                var result8 = System.Math.Round((double)sheet.Cells["B11"].Value, 0);
                var result9 = System.Math.Round((double)sheet.Cells["A12"].Value, 1);
                var result10 = System.Math.Round((double)sheet.Cells["B12"].Value, 1);
                Assert.AreEqual(1.4d, result1);
                Assert.AreEqual(3.4d, result2);
                Assert.AreEqual(1.587450787d, result3);
                Assert.AreEqual(2.969848481d, result4);
                Assert.AreEqual(0.28d, result5);
                Assert.AreEqual(3.54964787d, result6);
                Assert.AreEqual(0.777777778d, result7);
                Assert.AreEqual(2d, result8);
                Assert.AreEqual(9.8d, result9);
                Assert.AreEqual(25.2, result10);
            }
        }
    }
}
