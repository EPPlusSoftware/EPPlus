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

        [TestMethod]
        public void TroTrailingEmpty()
        {
            using(var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[3];
                sheet.Cells["B17"].Formula = "TROTRAILING(A18:G25)";
                sheet.Cells["B17"].Calculate();
                Assert.AreEqual(ErrorValues.RefError, sheet.Cells["B17"].Value);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TroAll3()
        {
            using (var package = OpenTemplatePackage("Trimfunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[3];
                sheet.Cells["H12"].Formula = "TROALL(A1:F5)";
                sheet.Cells["H12"].Calculate();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TrimRange()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets[4];
                sheet.Cells["I6"].Formula = "TRIMRANGE(A1:F5, 0, 0)";

                sheet.Cells["I12"].Formula = "TRIMRANGE(A1:F5, 1, 1)";

                sheet.Cells["I17"].Formula = "TRIMRANGE(A1:F5, 2, 2)"; 

                sheet.Cells["I22"].Formula = "TRIMRANGE(A1:F5, 3, 3)";
                sheet.Calculate();

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TrimRange_WithArrayConstantAsRange_TrimsCorrectly()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[5];

                // Array-konstant med tomma celler i kanterna
                sheet.Cells["G2"].Formula = "TRIMRANGE({\"\",\"\",\"\",\"\",1,2,\"\",3,4})";
                sheet.Cells["G2"].Calculate();
                   
                // Förväntad output: 2x2-array med 1,2,3,4
                SaveAndCleanup(package);
                //Assert.AreEqual(1d, sheet.Cells["A5"].Value);
                //Assert.AreEqual(2d, sheet.Cells["A6"].Value);
                //Assert.AreEqual(3d, sheet.Cells["A8"].Value);
                //Assert.AreEqual(4d, sheet.Cells["A9"].Value);
            }
        }

    }
}
