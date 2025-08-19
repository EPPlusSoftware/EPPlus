using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class OffsetTests
    {
        [TestMethod]
        public void OffsetShouldSetDefaultHeightIfOmitted()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 2;
            ws.Cells["A2"].Value = 3;
            ws.Cells["B2"].Value = 4;

            ws.Cells["D3"].Formula = "OFFSET(A1:B2,0,0,,2)";
            ws.Calculate();

            Assert.AreEqual(1, ws.Cells["D3"].Value);
            Assert.AreEqual(2, ws.Cells["E3"].Value);
            Assert.AreEqual(3, ws.Cells["D4"].Value);
            Assert.AreEqual(4, ws.Cells["E4"].Value);
        }

        [TestMethod]
        public void OffsetShouldSetDefaultWidthIfOmitted()
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 2;
            ws.Cells["A2"].Value = 3;
            ws.Cells["B2"].Value = 4;

            ws.Cells["D3"].Formula = "OFFSET(A1:B2,0,0,2,)";
            ws.Calculate();

            Assert.AreEqual(1, ws.Cells["D3"].Value);
            Assert.AreEqual(2, ws.Cells["E3"].Value);
            Assert.AreEqual(3, ws.Cells["D4"].Value);
            Assert.AreEqual(4, ws.Cells["E4"].Value);
        }
    }
}
