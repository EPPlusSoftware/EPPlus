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
    public class ImCoshTests
    {
        [TestMethod]
        public void ImCoshShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCOSH(\"4+3i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-27.0349456030742+3.85115333481178i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCoshShouldReturnCorrectNegativeResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCOSH(\"4-3i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("-27.0349456030742-3.85115333481178i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCoshShouldReturnOnlyiOrjResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCOSH(\"i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.54030230586814", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCoshShouldReturnOnlyImagValue()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCOSH(\"0+i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.54030230586814", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCoshShouldReturnOnlyHighValue()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCOSH(\"170+58000i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
#if Core
                Assert.AreEqual("3.3690976136249E+73-2.8221472985184E+72i", result); //Small diffrence in the result compared to Excel beacause .NET Core uses more decimals.
#else
                Assert.AreEqual("3.3690976136249E+73-2.82214729851839E+72i", result); //Small diffrence in the result compared to Excel beacause .NET Framework uses more decimals.
#endif
            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImCoshShouldReturnErrrorNumWhenInfinateResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMCOSH(\"17000+5000000i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ErrorValues.NumError, result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }
    }
}