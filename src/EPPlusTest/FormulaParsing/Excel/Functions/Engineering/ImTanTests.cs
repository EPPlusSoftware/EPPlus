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
    public class ImTanTests
    {

        [TestMethod]
        public void ImTanShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMTAN(\"4+3i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.00490825806749606+1.00070953606723i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImTanShouldReturnCorrectiValue()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMTAN(\"-i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-0.761594155955765i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImTanShouldReturnCorrectResultDecimalInput()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMTAN(\"22.46+34.567i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("1.52390779750415E-30+i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImTanShouldReturnCorrectiResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMTAN(\"200\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-1.79252748379038", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;

        }
    }

}
