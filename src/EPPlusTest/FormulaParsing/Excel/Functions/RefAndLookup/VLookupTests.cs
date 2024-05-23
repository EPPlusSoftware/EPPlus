using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class VLookupTests : TestBase
    {
        [TestMethod]
        public void VlookupShouldHandleWholeColumn()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["D1"].Value = 1;
                sheet.Cells["D2"].Value = 2;
                sheet.Cells["D3"].Value = 2;
                sheet.Cells["D4"].Value = 3;
                sheet.Cells["D5"].Value = 3;
                sheet.Cells["D6"].Value = 4;
                sheet.Cells["D7"].Value = 4;
                sheet.Cells["D8"].Value = 5;
                sheet.Cells["D9"].Value = 5;

                sheet.Cells["E1"].Value = "a";
                sheet.Cells["E2"].Value = "b";
                sheet.Cells["E3"].Value = "c";
                sheet.Cells["E4"].Value = "d";
                sheet.Cells["E5"].Value = "e";
                sheet.Cells["E6"].Value = "f";
                sheet.Cells["E7"].Value = "g";
                sheet.Cells["E8"].Value = "h";
                sheet.Cells["E9"].Value = "i";

                sheet.Cells["C10"].Formula = "VLOOKUP(3,D:E,2,FALSE)";
                sheet.Calculate();
                Assert.AreEqual("d", sheet.Cells["C10"].Value);
            }
        }
    }
}
