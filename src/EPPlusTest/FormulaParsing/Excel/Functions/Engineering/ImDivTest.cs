using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ImDivTest
    {
        [TestMethod]
        public void ImDivShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMDIV(\"-238+240i\", \"10+24i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("5+12i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImDivShouldReturnCorrectResult2()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMDIV(\"9+3i\", \"6\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("1.5+0.5i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImDivShouldReturnCorrectResult3()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMDIV(\"2+2i\", \"2+i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("1.2+0.4i", result);
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }
    }
}
