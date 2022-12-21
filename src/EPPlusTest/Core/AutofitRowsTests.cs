using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class AutofitRowsTests : TestBase
    {

        [TestMethod]
        public void AutofitRow_ShouldCalculateNewRowHeightWhenWrapTextIsTrue()
        {
            using (var pck = OpenPackage("AutofitRows_CustomWidth_WrapTextTrue.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["A1"].Style.WrapText = true;
                sheet.Cells["A1"].AutoFitRows();
                Assert.AreEqual(122.2d, sheet.Row(1).Height);
                SaveAndCleanup(pck);
            }
        }

        [TestMethod]
        public void AutofitRow_ShouldNotCalculateNewRowHeightWhenWrapTextIsTrue()
        {
            var defaultWidth = 15d;
            using (var pck = OpenPackage("AutofitRows_CustomWidth_WrapText_False.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = "A long text that needs some serious autofit of row height";
                sheet.Cells["A1"].Style.WrapText = false;
                sheet.Cells["A1"].AutoFitRows();
                Assert.AreEqual(defaultWidth, sheet.Row(1).Height);
                SaveAndCleanup(pck);
            }
        }
    }
}
