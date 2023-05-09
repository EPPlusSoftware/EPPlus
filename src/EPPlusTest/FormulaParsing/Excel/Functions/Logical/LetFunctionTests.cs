using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class LetFunctionTests
    {
        [TestMethod]
        public void Test1()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "LET(x,1 * 4, x + 1)";
                sheet.Calculate();
                Assert.AreEqual(2, sheet.Cells["A1"].Value);
            }
        }
    }
}
