using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.TextFunctions
{
    [TestClass]
    public class ReplaceTests
    {
        [TestMethod]
        public void ReplaceTest1()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Formula = "REPLACE(\"OLDTEXT\",5,20,\"NEW\")";
                sheet.Calculate();
                Assert.AreEqual("OLDTNEW", sheet.Cells["A1"].Value);
            }
        }
    }
}
