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
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets["TROALL1"];
                sheet.Cells["B26"].Formula = "_TRO_ALL(A1:E8)";
                sheet.Cells["B26"].Calculate();

                Assert.AreEqual("A", sheet.Cells["B27"].Value);
                Assert.AreEqual("TROTRAILING", sheet.Cells["D32"].Value);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TroAll2()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode= ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets["TROALL2"];
                sheet.Cells["B10"].Formula = "_TRO_ALL(A1:F5)";
                sheet.Cells["B10"].Calculate();
                Assert.AreEqual("A", sheet.Cells["B10"].Value);
                Assert.AreEqual("A", sheet.Cells["E12"].Value);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TroLeading()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets["TROLEADING1"];
                sheet.Cells["B12"].Formula = "_TRO_LEADING(A1:F5)";
                sheet.Cells["B12"].Calculate();

                Assert.AreEqual("A", sheet.Cells["B12"].Value);
                Assert.AreEqual(0d, sheet.Cells["F15"].Value);
                SaveAndCleanup(package);
            } 
        }

        [TestMethod]
        public void TroLeadingEmpty()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets["TROLEADING1"];
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
                var sheet = package.Workbook.Worksheets["TROTRAILING1"];
                sheet.Cells["B12"].Formula = "_TRO_TRAILING(A1:G5)";
                sheet.Calculate();
                Assert.AreEqual(0d, sheet.Cells["B12"].Value);
                Assert.AreEqual("A", sheet.Cells["F15"].Value);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TroTrailingEmpty()
        {
            using(var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets["TROTRAILING1"];
                sheet.Cells["B17"].Formula = "_TRO_TRAILING(A18:G25)";
                sheet.Cells["B17"].Calculate();
                Assert.AreEqual(ErrorValues.RefError, sheet.Cells["B17"].Value);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void TrimRangeColsTrailingOnly()
        {
            // Verifies that argument 3 (trim_cols) trims columns, not rows.
            // TRIMRANGE(A1:F5, 0, 2):
            //   trim_rows=None     -> keeps all 5 rows
            //   trim_cols=Trailing -> removes trailing empty col F
            using (var package = new ExcelPackage())
            {
                // Arrange: 5 rows x 6 cols, with empty leading/trailing rows and an empty trailing col F
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["B2"].Value = "A";
                sheet.Cells["C2"].Value = "A";
                sheet.Cells["E2"].Value = "A";
                sheet.Cells["B3"].Value = "A";
                sheet.Cells["C3"].Value = "A";
                sheet.Cells["E3"].Value = "A";
                sheet.Cells["B4"].Value = "A";
                sheet.Cells["C4"].Value = "A";
                sheet.Cells["E4"].Value = "A";
                // Rows 1 and 5 are empty. Column F is empty. A and D are empty.

                package.Workbook.CalcMode = ExcelCalcMode.Manual;

                // Act
                sheet.Cells["I10"].Formula = "TRIMRANGE(A1:F5, 0, 2)";
                sheet.Calculate();

                // Assert: 5 rows x 5 cols starting at I10, F column was trimmed
                // Expected Excel output:
                //   0  0  0  0  0
                //   0  A  A  0  A
                //   0  A  A  0  A
                //   0  A  A  0  A
                //   0  0  0  0  0

                // Last column of result (M) must exist - it is the original col E
                Assert.AreEqual("A", sheet.Cells["M11"].Value);
                Assert.AreEqual("A", sheet.Cells["M12"].Value);
                Assert.AreEqual("A", sheet.Cells["M13"].Value);

                // The 6th column would land at N - it must NOT exist in the result
                // (if trim_rows and trim_cols were swapped, F would still be there)
                Assert.IsNull(sheet.Cells["N10"].Value);

                // Last row of result (row 14) must exist - no row trim was requested
                // (if args were swapped, fewer rows would remain)
                Assert.IsNotNull(sheet.Cells["I14"].Value);

                // Spot-check the data row
                Assert.AreEqual("A", sheet.Cells["J11"].Value);
                Assert.AreEqual("A", sheet.Cells["K11"].Value);
            }
        }

        [TestMethod]
        public void TrimRangeRowsTrailingOnly()
        {
            // Verifies that argument 2 (trim_rows) trims rows, not columns.
            // TRIMRANGE(A1:F5, 2, 0):
            //   trim_rows=Trailing -> removes trailing empty row 5
            //   trim_cols=None     -> keeps all 6 cols
            using (var package = new ExcelPackage())
            {
                // Arrange: same layout as above
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["B2"].Value = "A";
                sheet.Cells["C2"].Value = "A";
                sheet.Cells["E2"].Value = "A";
                sheet.Cells["B3"].Value = "A";
                sheet.Cells["C3"].Value = "A";
                sheet.Cells["E3"].Value = "A";
                sheet.Cells["B4"].Value = "A";
                sheet.Cells["C4"].Value = "A";
                sheet.Cells["E4"].Value = "A";

                package.Workbook.CalcMode = ExcelCalcMode.Manual;

                // Act
                sheet.Cells["I10"].Formula = "TRIMRANGE(A1:F5, 2, 0)";
                sheet.Calculate();

                // Assert: row 5 trimmed, all 6 cols kept -> 4 rows x 6 cols at I10:N13

                // 6th column of result (N) must exist - no col trim was requested
                // (if args were swapped, col F would have been trimmed)
                Assert.IsNotNull(sheet.Cells["N10"].Value);

                // 5th row of result (row 14) must NOT exist - row 5 was trimmed
                // (if args were swapped, all 5 rows would remain)
                Assert.IsNull(sheet.Cells["I14"].Value);

                // Spot-check the data
                Assert.AreEqual("A", sheet.Cells["J11"].Value);
                Assert.AreEqual("A", sheet.Cells["K11"].Value);
                Assert.AreEqual("A", sheet.Cells["M11"].Value);
            }
        }

        [TestMethod]
        public void TrimRange()
        {
            using (var package = OpenTemplatePackage("TrimFunctions.xlsx"))
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var sheet = package.Workbook.Worksheets["TRIMRANGE"];
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
