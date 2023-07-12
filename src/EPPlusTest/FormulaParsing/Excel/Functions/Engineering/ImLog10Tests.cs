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
    public class ImLog10Test
    {
        [TestMethod]
        public void ImLog10ShouldReturnCorrectResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMLOG10(\"3+4i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.698970004336019+0.402719196273373i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImLog10ShouldReturniResult()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMLOG10(\"i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.682188176920921i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImLog10ShouldReturnCorrectResultBigInput()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMLOG10(\"12777+67423i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("4.83646954018446+0.600851676733113i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }

        [TestMethod]
        public void ImLog10ShouldReturnCorrectResultHighDecimalInput()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            using (var package = new ExcelPackage())

            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMLOG10(\"0+i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("0.682188176920921i", result);

            }
            Thread.CurrentThread.CurrentCulture = ci;
        }
    }
}

