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
    }
}
