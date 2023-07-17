using EPPlusTest.ThreadedComments;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ImPowerTest
    {
        [TestMethod]
        public void ImPowerShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(\"2+3i\",3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-46+9.00000000000001i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }


        [TestMethod]
        public void ImPowerShouldReturnCorrectresultWithIOnlyInput()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(\"0+3i\",3)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-4.96185124237991E-15-27i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImPowerShouldReturnCorrectResult2()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(\"5+12i\", 2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-119+120i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }
        [TestMethod]
        public void ImPowerShouldReturnCorrectResultRealValueOnly()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(5,2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("25", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImPowerShouldReturnCorrectResultHighPower()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(\"4-3i\",100)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("4.14265194823336E+68-7.87772410833036E+69i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod, Ignore]
        public void ImPowerShouldReturnCorrectresultWithIOnlyInput2()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(\"i\", 2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-1+1.22514845490862E-16i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImPowerShouldReturnCorrectResultDecimalInput()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMPOWER(\"4.4332-3.46554i\",4)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-885.720207814241-469.708954290734i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

    }
}