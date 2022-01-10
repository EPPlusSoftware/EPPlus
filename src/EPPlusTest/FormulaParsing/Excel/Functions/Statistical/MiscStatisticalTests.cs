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
    public class MiscStatisticalTests
    {
        [TestMethod]
        public void CorrelTest1()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 3;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 4;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 6;
                sheet.Cells["B1"].Value = 9;
                sheet.Cells["B2"].Value = 7;
                sheet.Cells["B3"].Value = 12;
                sheet.Cells["B4"].Value = 15;
                sheet.Cells["B5"].Value = 17;
                sheet.Cells["B6"].Formula = "CORREL(A1:A5,B1:B5)";
                sheet.Calculate();
                var result = sheet.Cells["B6"].Value;

                Assert.AreEqual(0.997054, System.Math.Round((double)result, 6));
            }
        }

        [TestMethod]
        public void FisherTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 0.75;
                sheet.Cells["A2"].Formula = "FISHER(A1)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;

                Assert.AreEqual(0.9729551, System.Math.Round((double)result, 7));
            }
        }
    }
}
