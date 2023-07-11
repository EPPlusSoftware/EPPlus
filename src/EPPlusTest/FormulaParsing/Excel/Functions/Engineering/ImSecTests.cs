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
    public class ImSecTest
    {
        [TestMethod]
        public void ImSecShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSEC(\"4+3i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-0.0652940278579471-0.0752249603027732i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImSecShouldReturniResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSEC(\"i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.648054273663885", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImSecShouldCorrectHighDecimalInputResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSEC(\"4.12321+3.7654745i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-0.0257596371683265-0.0384998326861422i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImSecShouldReturnZero()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSEC(\"17000+5000000i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImSecShouldReturnImaginaryPart()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSEC(\"0+i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.648054273663885", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImSecShouldReturnCorrectResult2()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMSEC(\"4-3i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-0.0652940278579471+0.0752249603027732i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }
    }
}