using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using M = System.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class NormInvTests
    {
        [TestMethod]
        public void NormInvShouldReturnCorrectResult()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "NORM.INV(0.25,0,1)";
                sheet.Calculate();
                Assert.AreEqual(-0.67449, M.Round(sheet.Cells["A1"].GetValue<double>(), 5));

                sheet.Cells["A2"].Formula = "NORMINV(0.25,0,1)";
                sheet.Calculate();
                Assert.AreEqual(-0.67449, M.Round(sheet.Cells["A2"].GetValue<double>(), 5));

                sheet.Cells["A3"].Formula = "NORMINV(0.84134,3,2)";
                sheet.Calculate();
                Assert.AreEqual(4.99996, M.Round(sheet.Cells["A3"].GetValue<double>(), 5));

                sheet.Cells["A4"].Formula = "NORM.INV(0.6589453152,2.58478,0.888)";
                sheet.Calculate();
                Assert.AreEqual(2.94849, M.Round(sheet.Cells["A4"].GetValue<double>(), 5));
            }
        }

        [TestMethod]
        public void NormsInvShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "NORM.S.INV(0.84134)";
                sheet.Calculate();
                Assert.AreEqual(0.99998, M.Round(sheet.Cells["A1"].GetValue<double>(), 5));

                sheet.Cells["A2"].Formula = "NORMSINV(0.995)";
                sheet.Calculate();
                Assert.AreEqual(2.57583, M.Round(sheet.Cells["A2"].GetValue<double>(), 5));
            }
        }
    }
}
