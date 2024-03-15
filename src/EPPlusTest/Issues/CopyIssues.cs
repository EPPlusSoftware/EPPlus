using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class CopyIssues : TestBase
    {
        [TestMethod]
        public void Issue1332()
        {
            // the error in this issue was that the intersect operator (SPACE)
            // was replaced with "isc" when a formulas was copied to a new destination
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "SUBTOTAL(109, _DATA _Quantity)";
            sheet.Cells["A1"].Copy(sheet.Cells["B1"]);
            Assert.AreEqual("SUBTOTAL(109,_DATA _Quantity)", sheet.Cells["B1"].Formula);
        }
    }
}
