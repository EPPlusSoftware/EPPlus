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
    public class TrimFunctionsTests : TestBase
    {
        [TestMethod]
        public void TroAll()
        {
            using (var package = OpenTemplatePackage("Trimfunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[0];
                sheet.Cells["B26"].Formula = "TROALL(A1:E8)";
                sheet.Cells["B26"].Calculate();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TroAll2()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode= ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[1];
                sheet.Cells["B10"].Formula = "TROALL(A1:F5)";
                sheet.Cells["B10"].Calculate();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TroLeading()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[2];
                sheet.Cells["B12"].Formula = "TROLEADING(A1:F5)";
                sheet.Cells["B12"].Calculate();

                SaveAndCleanup(package);
            } 
        }

        [TestMethod]
        public void TroLeadingEmpty()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[2];
                sheet.Cells["B18"].Formula = "TROLEADING(A19:E25)";
                sheet.Cells["B18"].Calculate();
                Assert.AreEqual(ErrorValues.RefError, sheet.Cells["B18"].Value);                
            }
        }

        [TestMethod]
        public void TroTrailing()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[3];
                sheet.Cells["B12"].Formula = "TROTRAILING(A1:G5)";
                sheet.Calculate();
                SaveAndCleanup(package);
            }
        }
    }
}
