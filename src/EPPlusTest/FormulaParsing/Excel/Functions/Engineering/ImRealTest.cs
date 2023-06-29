using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering
{
    [TestClass]
    public class ImRealTest
    {
        [TestMethod]
        public void ImRealShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMREAL(\"3+5i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(3d, result);
            }
        }

        [TestMethod]
        public void ImRealShouldReturnCorrectDataType()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMREAL(\"3+5i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.IsInstanceOfType(result, typeof(double));
            }
        }

        [TestMethod]
        public void ImRealShouldReturnWhenRealIsDecimalNumber()
        {
            var cc = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Formula = "IMREAL(\"3.5+5i\")";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(3.5d, result);
            }
            Thread.CurrentThread.CurrentCulture = cc;
        }
    }
}
