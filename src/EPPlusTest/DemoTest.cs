using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class DemoTest : TestBase
    {
        [TestMethod]
        public void Demo1()
        {
            using (var p = OpenTemplatePackage("TwrReturn.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                p.Workbook.Calculate();
                Assert.AreEqual(0.0677903, Math.Round((double)ws.Cells["E2"].Value, 7));
                Assert.AreEqual(0.0013835, Math.Round((double)ws.Cells["E3"].Value, 7));
            }
        }
        [TestMethod]
        public void Demo2()
        {
            using (var p = OpenTemplatePackage("TwrReturn.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                ws.Calculate();

                Assert.AreEqual(-0.2499483, Math.Round((double)ws.Cells["A1"].Value,7));
            }
        }
    }
}
