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
    public class ComplexTests
    {
        [TestMethod]
        public void ComplexShouldReturnCorrectResult()
        {
            var comma = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                
                sheet.Cells["A1"].Formula = "COMPLEX(5,2)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual("5+2i", result);

                sheet.Cells["A1"].Formula = "COMPLEX(5,-2.5, \"j\")";
                sheet.Calculate();
                result = sheet.Cells["A1"].Value;
                Assert.AreEqual($"5-2{comma}5j", result);
            }
        }
    }
}
