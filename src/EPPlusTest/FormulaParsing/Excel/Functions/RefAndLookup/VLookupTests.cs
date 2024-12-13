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
        [TestMethod]
        public void PriorAddressExpressionWorksheetShouldBeCleared()
        {
            using (var pck = OpenPackage("vlookuptest.xlsx", true))
            {
                #region firstWorksheet
                using var firstWorksheet = pck.Workbook.Worksheets.Add("firstWorksheet");

                firstWorksheet.SetValue("A1", 4000);
                firstWorksheet.Names.Add("search", new ExcelRange(firstWorksheet, "A1"));

                firstWorksheet.SetValue("B53", 0); firstWorksheet.SetValue("C53", -1); firstWorksheet.SetValue("D53", -1);
                firstWorksheet.SetValue("B54", 3500); firstWorksheet.SetValue("C54", -1); firstWorksheet.SetValue("D54", 151);
                firstWorksheet.SetValue("B55", 4500); firstWorksheet.SetValue("C55", -1); firstWorksheet.SetValue("D55", -1);

                firstWorksheet.SetFormula(2, 1, "VLOOKUP(firstWorksheet!search,$B$53:$D$55,3,1)");

                pck.Workbook.Calculate();

                Assert.AreEqual(151, firstWorksheet.Cells["A2"].Value);
                #endregion

                #region secondWorksheet
                using var secondWorksheet = pck.Workbook.Worksheets.Add("secondWorksheet");

                secondWorksheet.SetValue("B53", 0); secondWorksheet.SetValue("C53", -1); secondWorksheet.SetValue("D53", -1);
                secondWorksheet.SetValue("B54", 3500); secondWorksheet.SetValue("C54", -1); secondWorksheet.SetValue("D54", 251);
                secondWorksheet.SetValue("B55", 4500); secondWorksheet.SetValue("C55", -1); secondWorksheet.SetValue("D55", -1);

                secondWorksheet.SetFormula(2, 1, "VLOOKUP(firstWorksheet!search,$B$53:$D$55,3,1)");

                secondWorksheet.Calculate();

                Assert.AreEqual(251, secondWorksheet.Cells["A2"].Value);
                SaveAndCleanup(pck);
                #endregion
            }
        }
    }
}
