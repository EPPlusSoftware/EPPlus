using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class SubtotalAutoAndTableFilterTests : TestBase
    {
        [TestMethod]
        public void TestWorkbook()
        {
            using(var package = OpenTemplatePackage("SubtotalFilters.xlsx"))
            {
                package.Workbook.Calculate();

                var sheet1 = package.Workbook.Worksheets["Default"];
                Assert.AreEqual(6d, sheet1.Cells["B6"].Value, "Default B6");
                Assert.AreEqual(6d, sheet1.Cells["B7"].Value, "Default B7");
                Assert.AreEqual(6d, sheet1.Cells["B18"].Value, "Default B18");
                Assert.AreEqual(6d, sheet1.Cells["B19"].Value, "Default B19");

                var sheet2 = package.Workbook.Worksheets["HiddenRowNoFilterNoTable"];
                Assert.AreEqual(6d, sheet2.Cells["B6"].Value, "HiddenRowNoFilterNoTable B6");
                Assert.AreEqual(3d, sheet2.Cells["B7"].Value, "HiddenRowNoFilterNoTable B7");
                Assert.AreEqual(6d, sheet2.Cells["B18"].Value, "HiddenRowNoFilterNoTable B18");
                Assert.AreEqual(3d, sheet2.Cells["B19"].Value, "HiddenRowNoFilterNoTable B19");

                var sheet3 = package.Workbook.Worksheets["HiddenRowWithAutoFilterNoFilter"];
                Assert.AreEqual(6d, sheet3.Cells["B6"].Value, "HiddenRowWithAutoFilterNoFilter B6");
                Assert.AreEqual(3d, sheet3.Cells["B7"].Value, "HiddenRowWithAutoFilterNoFilter B7");
                Assert.AreEqual(6d, sheet3.Cells["B18"].Value, "HiddenRowWithAutoFilterNoFilter B18");
                Assert.AreEqual(3d, sheet3.Cells["B19"].Value, "HiddenRowWithAutoFilterNoFilter B19");

                var sheet4 = package.Workbook.Worksheets["HiddenRowWithTableNoFilter"];
                Assert.AreEqual(6d, sheet4.Cells["B6"].Value, "HiddenRowWithTableNoFilter B6");
                Assert.AreEqual(3d, sheet4.Cells["B7"].Value, "HiddenRowWithTableNoFilter B7");
                Assert.AreEqual(6d, sheet4.Cells["B18"].Value, "HiddenRowWithTableNoFilter B18");
                Assert.AreEqual(3d, sheet4.Cells["B19"].Value, "HiddenRowWithTableNoFilter B19");

                var sheet5 = package.Workbook.Worksheets["HiddenByAutoFilter"];
                Assert.AreEqual(3d, sheet5.Cells["B6"].Value, "HiddenByAutoFilter B6");
                Assert.AreEqual(3d, sheet5.Cells["B7"].Value, "HiddenByAutoFilter B7");
                Assert.AreEqual(3d, sheet5.Cells["B18"].Value, "HiddenByAutoFilter B18");
                Assert.AreEqual(3d, sheet5.Cells["B19"].Value, "HiddenByAutoFilter B19");

                var sheet6 = package.Workbook.Worksheets["HiddenByAutoFilter"];
                Assert.AreEqual(3d, sheet6.Cells["B6"].Value, "HiddenByTableFilter B6");
                Assert.AreEqual(3d, sheet6.Cells["B7"].Value, "HiddenByTableFilter B7");
                Assert.AreEqual(3d, sheet6.Cells["B18"].Value, "HiddenByTableFilter B18");
                Assert.AreEqual(3d, sheet6.Cells["B19"].Value, "HiddenByTableFilter B19");
            }
        }
    }
}
