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
    public class ImCschTest
    {
        [TestMethod]
        public void ImCschShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCSCH(\"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-0.0648774713706355+0.0754898329158637i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCschShouldReturnCorrectResult2()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCSCH(\"4+3i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-0.036275889628626-0.0051744731840194i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCschShouldReturniOnlyResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCSCH(\"-i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("1.18839510577812i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCschShouldReturnCorrectZeroIfNullResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCSCH(\"4-3410.34343453545i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.00527360904066106-0.0362377851539404i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCschShouldReturnCorrectWithDecimalInput()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCSCH(\"33.34543+8.21312i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-2.31835090849203E-15-6.17536156326111E-15i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCrchShouldReturnCorrectResultOnlyRealNumber()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCSCH(\"200\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("2.76779305347348E-87", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;

        }
    }
}