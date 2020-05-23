using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ErfTests
    {
        [TestMethod]
        public void ErfTest1()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                
                sheet.Cells["A1"].Formula = "ERF(1.5)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.966105146, System.Math.Round((double)result, 9));

                sheet.Cells["A1"].Formula = "ERF(0, 1.5)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.966105146, System.Math.Round((double)result, 9));

                sheet.Cells["A1"].Formula = "ERF(1, 2)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.152621472, System.Math.Round((double)result, 9));

            }
        }

        [TestMethod]
        public void ErfPreciseTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Formula = "ERF.PRECISE(-1)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(-0.842700793, System.Math.Round((double)result, 9));

                sheet.Cells["A1"].Formula = "ERF.PRECISE(0.745)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.70792892, System.Math.Round((double)result, 8));

                sheet.Cells["A1"].Formula = "ERF.PRECISE(1.5)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.966105146, System.Math.Round((double)result, 9));

            }
        }

        [TestMethod]
        public void ErfcTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");

                sheet.Cells["A1"].Formula = "ERFC(0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(1d, System.Math.Round((double)result, 9));

                sheet.Cells["A1"].Formula = "ERFC(0.5)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(0.479500122, System.Math.Round((double)result, 9));

                sheet.Cells["A1"].Formula = "ERFC(-1)";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual(1.842700793, System.Math.Round((double)result, 9));

            }
        }
    }
}
