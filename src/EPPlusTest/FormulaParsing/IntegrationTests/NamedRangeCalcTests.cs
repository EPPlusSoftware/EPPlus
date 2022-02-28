using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class NamedRangeCalcTests
    {
        [TestMethod]
        public void Sum_NamedRangeToNamedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var wks = pck.Workbook.Worksheets.Add("Sheet1");
                wks.Cells["C3"].Value = 7;
                wks.Cells["C4"].Value = 1;
                wks.Cells["C5"].Value = 2;
                wks.Names.Add("MyName1", wks.Cells["C3"]);
                wks.Names.Add("MyName2", wks.Cells["C5"]);
                wks.Cells["F6"].Formula = "SUM(C3:C5)";
                wks.Cells["F7"].Formula = "SUM(MyName1:MyName2)";

                pck.Workbook.Calculate();

                Assert.AreEqual(10, wks.Cells["F6"].GetValue<double>(), 1E-10);
                Assert.AreEqual(10, wks.Cells["F7"].GetValue<double>(), 1E-10);
            }
        }

        [TestMethod]
        public void SumProduct_NamedRangesToNamedRanges()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 10;
                sheet1.Cells["A2"].Value = 4;
                sheet1.Cells["A3"].Value = 1.6;
                sheet1.Cells["B1"].Value = 5;
                sheet1.Cells["B2"].Value = 7;
                sheet1.Cells["B3"].Value = 10;
                sheet1.Cells["C1"].Value = 1;
                sheet1.Cells["C2"].Value = 0.5;
                sheet1.Cells["C3"].Value = 0.5;

                sheet1.Names.Add("MyName1", sheet1.Cells["A1"]);
                sheet1.Names.Add("MyName2", sheet1.Cells["A3"]);
                sheet1.Names.Add("MyName3", sheet1.Cells["B1"]);
                sheet1.Names.Add("MyName4", sheet1.Cells["B3"]);
                sheet1.Names.Add("MyName5", sheet1.Cells["C1"]);
                sheet1.Names.Add("MyName6", sheet1.Cells["C3"]);

                sheet1.Cells["C6"].Formula = "SUMPRODUCT(MyName1:MyName2,MyName3:MyName4,MyName5:MyName6)";

                pck.Workbook.Calculate();

                Assert.AreEqual(72, sheet1.Cells["C6"].GetValue<double>(), 0);
            }
        }
    }
}
