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
                sheet.Cells["B26"].Formula = "_TRO_ALL(A1:E8)";
                sheet.Cells["B26"].Calculate();

                Assert.AreEqual(sheet.Cells["B26"].Value, string.Empty);
                Assert.AreEqual(sheet.Cells["D32"].Value, "TROTRAILING");
                
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
                sheet.Cells["B10"].Formula = "_TRO_ALL(A1:F5)";
                sheet.Cells["B10"].Calculate();
                Assert.AreEqual(sheet.Cells["B10"].Value, "A");
                Assert.AreEqual(sheet.Cells["E12"].Value, "A");

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
                sheet.Cells["B12"].Formula = "_TRO_LEADING(A1:F5)";
                sheet.Cells["B12"].Calculate();

                Assert.AreEqual(sheet.Cells["B12"].Value, "A");
                Assert.AreEqual(sheet.Cells["F15"].Value, 0d);
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
                sheet.Cells["B18"].Formula = "_TRO_LEADING(A19:E25)";
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
                sheet.Cells["B12"].Formula = "_TRO_TRAILING(A1:G5)";
                sheet.Calculate();
                Assert.AreEqual(sheet.Cells["B12"].Value, 0d);
                Assert.AreEqual(sheet.Cells["F15"].Value, "A");
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
                sheet.Cells["B17"].Formula = "_TRO_TRAILING(A18:G25)";
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
                sheet.Cells["H12"].Formula = "_TRO_ALL(A1:F5)";
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

                Assert.AreEqual("A", sheet.Cells["J7"].Value);
                Assert.AreEqual("A", sheet.Cells["K7"].Value);
                Assert.AreEqual("A", sheet.Cells["M7"].Value);
                Assert.AreEqual("A", sheet.Cells["J8"].Value);
                Assert.AreEqual("A", sheet.Cells["K8"].Value);
                Assert.AreEqual("A", sheet.Cells["M8"].Value);
                Assert.AreEqual("A", sheet.Cells["J9"].Value);
                Assert.AreEqual("A", sheet.Cells["K9"].Value);
                Assert.AreEqual("A", sheet.Cells["M9"].Value);

                Assert.AreEqual("A", sheet.Cells["I12"].Value);
                Assert.AreEqual("A", sheet.Cells["J12"].Value);
                Assert.AreEqual("A", sheet.Cells["L12"].Value);
                Assert.AreEqual("A", sheet.Cells["I13"].Value);
                Assert.AreEqual("A", sheet.Cells["J13"].Value);
                Assert.AreEqual("A", sheet.Cells["L13"].Value);
                Assert.AreEqual("A", sheet.Cells["I14"].Value);
                Assert.AreEqual("A", sheet.Cells["J14"].Value);
                Assert.AreEqual("A", sheet.Cells["L14"].Value);

                Assert.AreEqual("A", sheet.Cells["J18"].Value);
                Assert.AreEqual("A", sheet.Cells["K18"].Value);
                Assert.AreEqual("A", sheet.Cells["M18"].Value);
                Assert.AreEqual("A", sheet.Cells["J19"].Value);
                Assert.AreEqual("A", sheet.Cells["K19"].Value);
                Assert.AreEqual("A", sheet.Cells["M19"].Value);
                Assert.AreEqual("A", sheet.Cells["J20"].Value);
                Assert.AreEqual("A", sheet.Cells["K20"].Value);
                Assert.AreEqual("A", sheet.Cells["M20"].Value);

                Assert.AreEqual("A", sheet.Cells["I22"].Value);
                Assert.AreEqual("A", sheet.Cells["J22"].Value);
                Assert.AreEqual("A", sheet.Cells["L22"].Value);
                Assert.AreEqual("A", sheet.Cells["I23"].Value);
                Assert.AreEqual("A", sheet.Cells["J23"].Value);
                Assert.AreEqual("A", sheet.Cells["L23"].Value);
                Assert.AreEqual("A", sheet.Cells["I24"].Value);
                Assert.AreEqual("A", sheet.Cells["J24"].Value);
                Assert.AreEqual("A", sheet.Cells["L24"].Value);
                SaveAndCleanup(package);
            }
        }
    }
}
