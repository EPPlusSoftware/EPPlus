using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class DbFunctionTests : TestBase
    {
        [TestMethod]
        public void DProductTest()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = "Price";
                ws.Cells["A2"].Value = ">14";
                ws.Cells["A3"].Value = "Price";
                ws.Cells["A4"].Value = 100;
                ws.Cells["A5"].Value = 13;
                ws.Cells["A6"].Value = 45;

                ws.Cells["A7"].Formula = "DPRODUCT(A3:A6, A3, A1:A2)";
                ws.Calculate();
                var result = ws.Cells["A7"].Value;
                Assert.AreEqual(4500d, result);
            }
        }

        [TestMethod]
        public void DStdev()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = "Price";
                ws.Cells["A2"].Value = ">14";
                ws.Cells["A3"].Value = "Price";
                ws.Cells["A4"].Value = 100;
                ws.Cells["A5"].Value = 13;
                ws.Cells["A6"].Value = 45;

                ws.Cells["A7"].Formula = "DSTDEV(A3:A6, A3, A1:A2)";
                ws.Calculate();
                var result = System.Math.Round((double)ws.Cells["A7"].Value, 8);
                Assert.AreEqual(38.89087297d, result);
            }
        }

        [TestMethod]
        public void DStdevp()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = "Price";
                ws.Cells["A2"].Value = ">14";
                ws.Cells["A3"].Value = "Price";
                ws.Cells["A4"].Value = 100;
                ws.Cells["A5"].Value = 13;
                ws.Cells["A6"].Value = 45;

                ws.Cells["A7"].Formula = "DSTDEVP(A3:A6, A3, A1:A2)";
                ws.Calculate();
                var result = System.Math.Round((double)ws.Cells["A7"].Value, 1);
                Assert.AreEqual(27.5, result);
            }
        }
    }
}
