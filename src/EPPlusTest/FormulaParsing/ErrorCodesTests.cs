using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ErrorCodesTests
    {
        [TestMethod]
        public void ShouldSetDiv0InFunctions()
        {
            using(var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].Formula = "ROUND(2.3 + 1/0, 2)";
                ws.Calculate();
                Assert.AreEqual("#DIV/0!", ws.Cells["A1"].Value.ToString());
            }
        }
    }
}
