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
    public class LognormDotInvTests
    {
        [TestMethod]
        public void LognormDotInvTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.039084;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["A4"].Value = 1.2;

                sheet.Cells["B2"].Formula = "LOGNORM.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B2"].Value, 7);
                Assert.AreEqual(4.0000252, result);
            }
        }

        [TestMethod]
        public void LognormDotInvTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.039084;
                sheet.Cells["B2"].Value = 0.056084;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["B3"].Value = 2.3;
                sheet.Cells["A4"].Value = 1.2;
                sheet.Cells["B4"].Value = 1.4;

                sheet.Cells["A5"].Formula = "LOGNORM.INV(A2:B2,A3:B3,A4:B4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["A5"].Value, 7);
                var result2 = System.Math.Round((double)sheet.Cells["B5"].Value, 7);

                Assert.AreEqual(4.0000252, result);
                Assert.AreEqual(1.0790349, result2);
            }
        }

        [TestMethod]
        public void LognormDotInvShouldReturnError()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.039084;
                sheet.Cells["B2"].Value = 0.056084;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["B3"].Value = 2.3;
                sheet.Cells["A4"].Value = 1.2;
                sheet.Cells["B4"].Value = 1.4;

                sheet.Cells["A5"].Formula = "LOGNORM.INV(A2:B3,A3:B4,A4:B4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["A5"].Value, 7);
                var result2 = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                var errorResult = sheet.Cells["A6"].Value;
                var errorResult2 = sheet.Cells["B6"].Value;

                Assert.AreEqual(4.0000252, result);
                Assert.AreEqual(1.0790349, result2);
                Assert.AreEqual(ErrorValues.NumError, errorResult);
                Assert.AreEqual(ErrorValues.NumError, errorResult2);
            }
        }

        [TestMethod]
        public void LognormDotInvTest3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.039084;
                sheet.Cells["B2"].Value = 0.056084;
                sheet.Cells["C2"].Value = 0.034212;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["B3"].Value = 2.3;
                sheet.Cells["A4"].Value = 1.2;
                sheet.Cells["B4"].Value = 1.4;

                sheet.Cells["A5"].Formula = "LOGNORM.INV(A2:C2,A3:B3,A4:B4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["A5"].Value, 7);
                var result2 = System.Math.Round((double)sheet.Cells["B5"].Value, 7);
                var errorResult = sheet.Cells["C5"].Value;
                

                Assert.AreEqual(4.0000252, result);
                Assert.AreEqual(1.0790349, result2);
                Assert.AreEqual(ErrorValues.NAError, errorResult);
            }
        }
    }
}