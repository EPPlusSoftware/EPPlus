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
    public class PivotByTests : TestBase
    {
        [TestMethod]
        public void BasicPivotBy()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["C1"].Value = "Bertil";
                s.Cells["C2"].Value = "Joe";
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["D1"].Formula = "PIVOTBY(A1:A2,C1:C2, B1:B2, _xleta.SUM)";
                s.Calculate();
                s.Workbook.FullCalcOnLoad = false;
                s.Workbook.CalcMode = ExcelCalcMode.Manual;

                // Rubrikrad
                Assert.AreEqual("Bertil", s.Cells["E1"].Value);
                Assert.AreEqual("Joe", s.Cells["F1"].Value);
                Assert.AreEqual("Total", s.Cells["G1"].Value);

                // Anna-rad
                Assert.AreEqual("Anna", s.Cells["D2"].Value);
                Assert.AreEqual(2d, s.Cells["F2"].Value);
                Assert.AreEqual(2d, s.Cells["G2"].Value);

                // Joe-rad
                Assert.AreEqual("Joe", s.Cells["D3"].Value);
                Assert.AreEqual(1d, s.Cells["E3"].Value);
                Assert.AreEqual(1d, s.Cells["G3"].Value);

                // Total-rad
                Assert.AreEqual("Total", s.Cells["D4"].Value);
                Assert.AreEqual(1d, s.Cells["E4"].Value);
                Assert.AreEqual(2d, s.Cells["F4"].Value);
                Assert.AreEqual(3d, s.Cells["G4"].Value);

                SaveWorkbook("BasicPivotBy.xlsx", package);
            }
        }

        [TestMethod]
        public void PivotBy()
        {
            using (var package = new ExcelPackage())
            {
                SwitchToCulture("en-US");
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Stockholm";
                s.Cells["A2"].Value = "Linköping";
                s.Cells["A3"].Value = "Örebro";
                s.Cells["B1"].Value = 2026;
                s.Cells["B2"].Value = 2026;
                s.Cells["B3"].Value = 2025;
                s.Cells["C1"].Value = "Q2";
                s.Cells["C2"].Value = "Q1";
                s.Cells["C3"].Value = "Q2";
                s.Cells["D1"].Value = 34543;
                s.Cells["D2"].Value = 43265;
                s.Cells["D3"].Value = 75461;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:C3,D1:D3, _xleta.SUM)";
                s.Calculate();

                // Rubrikrad 1 (år)
                Assert.AreEqual(2025, s.Cells["F1"].Value);
                Assert.AreEqual(2026, s.Cells["G1"].Value);
                Assert.AreEqual(2026, s.Cells["H1"].Value);
                Assert.AreEqual("Total", s.Cells["I1"].Value);

                // Rubrikrad 2 (kvartal)
                Assert.AreEqual("Q2", s.Cells["F2"].Value);
                Assert.AreEqual("Q1", s.Cells["G2"].Value);
                Assert.AreEqual("Q2", s.Cells["H2"].Value);

                // Linköping
                Assert.AreEqual("Linköping", s.Cells["E3"].Value);
                Assert.AreEqual(43265d, s.Cells["G3"].Value);
                Assert.AreEqual(43265d, s.Cells["I3"].Value);

                // Örebro
                Assert.AreEqual("Örebro", s.Cells["E4"].Value);
                Assert.AreEqual(75461d, s.Cells["F4"].Value);
                Assert.AreEqual(75461d, s.Cells["I4"].Value);

                // Stockholm
                Assert.AreEqual("Stockholm", s.Cells["E5"].Value);
                Assert.AreEqual(34543d, s.Cells["H5"].Value);
                Assert.AreEqual(34543d, s.Cells["I5"].Value);

                // Total-rad
                Assert.AreEqual("Total", s.Cells["E6"].Value);
                Assert.AreEqual(75461d, s.Cells["F6"].Value);
                Assert.AreEqual(43265d, s.Cells["G6"].Value);
                Assert.AreEqual(34543d, s.Cells["H6"].Value);
                Assert.AreEqual(153269d, s.Cells["I6"].Value);
                SwitchBackToCurrentCulture();
            }
        }

        [TestMethod]
        public void PivotBySortOrder()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["A4"].Value = "B";
                s.Cells["A5"].Value = "C";
                s.Cells["A6"].Value = "C";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["B4"].Value = "Y";
                s.Cells["B5"].Value = "X";
                s.Cells["B6"].Value = "Y";
                s.Cells["C1"].Value = 2;
                s.Cells["C2"].Value = 4;
                s.Cells["C3"].Value = 1;
                s.Cells["C4"].Value = 5;
                s.Cells["C5"].Value = 7;
                s.Cells["C6"].Value = 4;
                s.Cells["D1"].Formula = "PIVOTBY(A1:A6,B1:B6,C1:C6,_xleta.SUM,,,-1,,-1)";
                s.Calculate();

                Assert.AreEqual("Y", s.Cells["E1"].Value);
                Assert.AreEqual("X", s.Cells["F1"].Value);
                Assert.AreEqual("Total", s.Cells["G1"].Value);

                // C
                Assert.AreEqual("C", s.Cells["D2"].Value);
                Assert.AreEqual(4d, s.Cells["E2"].Value);
                Assert.AreEqual(7d, s.Cells["F2"].Value);
                Assert.AreEqual(11d, s.Cells["G2"].Value);

                // B
                Assert.AreEqual("B", s.Cells["D3"].Value);
                Assert.AreEqual(5d, s.Cells["E3"].Value);
                Assert.AreEqual(1d, s.Cells["F3"].Value);
                Assert.AreEqual(6d, s.Cells["G3"].Value);

                // A
                Assert.AreEqual("A", s.Cells["D4"].Value);
                Assert.AreEqual(4d, s.Cells["E4"].Value);
                Assert.AreEqual(2d, s.Cells["F4"].Value);
                Assert.AreEqual(6d, s.Cells["G4"].Value);

                // Total
                Assert.AreEqual("Total", s.Cells["D5"].Value);
                Assert.AreEqual(13d, s.Cells["E5"].Value);
                Assert.AreEqual(10d, s.Cells["F5"].Value);
                Assert.AreEqual(23d, s.Cells["G5"].Value);
            }
        }

        [TestMethod]
        public void PivotBySubTotalsIncluded()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:C3,D1:D3,_xleta.SUM,,,,2)";
                s.Calculate();

                // Rubrikrad 1
                Assert.AreEqual("X", s.Cells["F1"].Value);
                Assert.AreEqual("X", s.Cells["G1"].Value);
                Assert.AreEqual("X", s.Cells["H1"].Value);
                Assert.AreEqual("Y", s.Cells["I1"].Value);
                Assert.AreEqual("Y", s.Cells["J1"].Value);
                Assert.AreEqual("Grand Total", s.Cells["K1"].Value);

                // Rubrikrad 2
                Assert.AreEqual("I", s.Cells["F2"].Value);
                Assert.AreEqual("O", s.Cells["G2"].Value);
                Assert.AreEqual("I", s.Cells["I2"].Value);

                // A
                Assert.AreEqual("A", s.Cells["E3"].Value);
                Assert.AreEqual(2d, s.Cells["G3"].Value);
                Assert.AreEqual(2d, s.Cells["H3"].Value);
                Assert.AreEqual(4d, s.Cells["I3"].Value);
                Assert.AreEqual(4d, s.Cells["J3"].Value);
                Assert.AreEqual(6d, s.Cells["K3"].Value);

                // B
                Assert.AreEqual("B", s.Cells["E4"].Value);
                Assert.AreEqual(1d, s.Cells["F4"].Value);
                Assert.AreEqual(1d, s.Cells["H4"].Value);
                Assert.AreEqual(1d, s.Cells["K4"].Value);

                // Total
                Assert.AreEqual("Total", s.Cells["E5"].Value);
                Assert.AreEqual(1d, s.Cells["F5"].Value);
                Assert.AreEqual(2d, s.Cells["G5"].Value);
                Assert.AreEqual(3d, s.Cells["H5"].Value);
                Assert.AreEqual(4d, s.Cells["I5"].Value);
                Assert.AreEqual(4d, s.Cells["J5"].Value);
                Assert.AreEqual(7d, s.Cells["K5"].Value);
            }
        }
        [TestMethod]
        public void PivotByGrandTotalsRows()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:B3,C1:C3,D1:D3,_xleta.PERCENTOF,,2)";
                s.Calculate();

                // Rubrikrad
                Assert.AreEqual("I", s.Cells["G1"].Value);
                Assert.AreEqual("O", s.Cells["H1"].Value);
                Assert.AreEqual("Total", s.Cells["I1"].Value);

                // A | X
                Assert.AreEqual("A", s.Cells["E2"].Value);
                Assert.AreEqual("X", s.Cells["F2"].Value);
                Assert.AreEqual(1d, s.Cells["H2"].Value);
                Assert.AreEqual(0.28571429, System.Math.Round((double)s.Cells["I2"].Value, 8));

                // A | Y
                Assert.AreEqual("A", s.Cells["E3"].Value);
                Assert.AreEqual("Y", s.Cells["F3"].Value);
                Assert.AreEqual(0.8d, s.Cells["G3"].Value);
                Assert.AreEqual(0.57142857, System.Math.Round((double)s.Cells["I3"].Value, 8));

                // A subtotal
                Assert.AreEqual("A", s.Cells["E4"].Value);
                Assert.AreEqual(0.8d, s.Cells["G4"].Value);
                Assert.AreEqual(1d, s.Cells["H4"].Value);
                Assert.AreEqual(0.85714286, System.Math.Round((double)s.Cells["I4"].Value, 8));

                // B | X
                Assert.AreEqual("B", s.Cells["E5"].Value);
                Assert.AreEqual("X", s.Cells["F5"].Value);
                Assert.AreEqual(0.2d, s.Cells["G5"].Value);
                Assert.AreEqual(0.14285714, System.Math.Round((double)s.Cells["I5"].Value, 8));

                // B subtotal
                Assert.AreEqual("B", s.Cells["E6"].Value);
                Assert.AreEqual(0.2d, s.Cells["G6"].Value);
                Assert.AreEqual(0.14285714, System.Math.Round((double)s.Cells["I6"].Value, 8));

                // Grand Total
                Assert.AreEqual("Grand Total", s.Cells["E7"].Value);
                Assert.AreEqual(1d, s.Cells["G7"].Value);
                Assert.AreEqual(1d, s.Cells["H7"].Value);
                Assert.AreEqual(1d, s.Cells["I7"].Value);
            }
        }

        [TestMethod]
        public void PivotByRelativeTo()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:B3,C1:C3,D1:D3,_xleta.PERCENTOF,,,,,,,3)";
                s.Calculate();
                
                Assert.AreEqual(0.714285714d, System.Math.Round((double)s.Cells["G5"].Value, 9));
                Assert.AreEqual(0.285714286d, System.Math.Round((double)s.Cells["H5"].Value, 9));
            }
        }

        [TestMethod]
        public void PivotByRelativeTo2()
        {
            using (var package = new ExcelPackage())
            {
                SwitchToCulture("en-US");
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Stockholm";
                s.Cells["A2"].Value = "Linköping";
                s.Cells["A3"].Value = "Örebro";
                s.Cells["A4"].Value = "Stockholm";
                s.Cells["A5"].Value = "Örebro";
                s.Cells["A6"].Value = "Linköping";

                s.Cells["B1"].Value = "2026";
                s.Cells["B2"].Value = "2026";
                s.Cells["B3"].Value = "2025";
                s.Cells["B4"].Value = "2025";
                s.Cells["B5"].Value = "2025";
                s.Cells["B6"].Value = "2024";

                s.Cells["C1"].Value = "Q2";
                s.Cells["C2"].Value = "Q1";
                s.Cells["C3"].Value = "Q2";
                s.Cells["C4"].Value = "Q3";
                s.Cells["C5"].Value = "Q4";
                s.Cells["C6"].Value = "Q2";

                s.Cells["D1"].Value = 34543;
                s.Cells["D2"].Value = 43265;
                s.Cells["D3"].Value = 75461;
                s.Cells["D4"].Value = 4536;
                s.Cells["D5"].Value = 64312;
                s.Cells["D6"].Value = 64531;

                s.Cells["E1"].Formula = "PIVOTBY(A1:A6,B1:C6,D1:D6,_xleta.PERCENTOF,,,,,,,3)";
                s.Calculate();
                SwitchBackToCurrentCulture();
            }
        }

        [TestMethod]
        public void PivotByRelativeTo3()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:C3,D1:D3,_xleta.PERCENTOF,,,,,,,4)";
                s.Calculate();
            }
        }

        [TestMethod]
        public void PivotByHeaders()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:B3,D1:D3, _xleta.SUM, 3)";
                s.Calculate();

                Assert.AreEqual("X", s.Cells["F2"].Value);
                Assert.AreEqual("A", s.Cells["E3"].Value);
                Assert.AreEqual("A", s.Cells["E4"].Value);
            }
        }

        [TestMethod]
        public void PivotByCustomLambdaWithHstack()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:B3,D1:D3, HSTACK(_xleta.COUNT, LAMBDA(x, SUM(x *2/3)), LAMBDA(x, SUM(x *2)) ),3)";
                s.Calculate();

            }
        }

        [TestMethod]
        public void PivotByCustomLambdaWithVstack()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:B3,D1:D3, VSTACK(_xleta.COUNT, LAMBDA(x, SUM(x *2/3)), LAMBDA(x, SUM(x *2)) ),3)";
                s.Calculate();
                Assert.AreNotEqual(0d, s.Cells["E1"].Value);
                Assert.AreNotEqual(0d, s.Cells["F1"].Value);
                Assert.AreNotEqual(0d, s.Cells["E2"].Value);
                Assert.AreNotEqual(0d, s.Cells["F2"].Value);

                // SaveWorkbook("PivotByCustomLambda.xlsx", package);
            }
        }

        [TestMethod]
        public void PivotBySortOrderArray()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = "O";
                s.Cells["C2"].Value = "I";
                s.Cells["C3"].Value = "I";
                s.Cells["D1"].Value = 2;
                s.Cells["D2"].Value = 4;
                s.Cells["D3"].Value = 1;
                s.Cells["E1"].Formula = "PIVOTBY(A1:B3,C1:C3,D1:D3, _xleta.SUM,,, {-1,-2})";
                s.Calculate();

                // Rubrikrad
                Assert.AreEqual("I", s.Cells["G1"].Value);
                Assert.AreEqual("O", s.Cells["H1"].Value);
                Assert.AreEqual("Total", s.Cells["I1"].Value);

                // B X
                Assert.AreEqual("B", s.Cells["E2"].Value);
                Assert.AreEqual("X", s.Cells["F2"].Value);
                Assert.AreEqual(1d, s.Cells["G2"].Value);
                Assert.AreEqual(1d, s.Cells["I2"].Value);

                // A Y
                Assert.AreEqual("A", s.Cells["E3"].Value);
                Assert.AreEqual("Y", s.Cells["F3"].Value);
                Assert.AreEqual(4d, s.Cells["G3"].Value);
                Assert.AreEqual(4d, s.Cells["I3"].Value);

                // A X
                Assert.AreEqual("A", s.Cells["E4"].Value);
                Assert.AreEqual("X", s.Cells["F4"].Value);
                Assert.AreEqual(2d, s.Cells["H4"].Value);
                Assert.AreEqual(2d, s.Cells["I4"].Value);

                // Total
                Assert.AreEqual("Total", s.Cells["E5"].Value);
                Assert.AreEqual(5d, s.Cells["G5"].Value);
                Assert.AreEqual(2d, s.Cells["H5"].Value);
                Assert.AreEqual(7d, s.Cells["I5"].Value);
            }
        }

        [TestMethod]

        public void PivotByTemplateTest()
        {
            using (var package = OpenTemplatePackage("PivotByTest1.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[1];

                sheet.Cells["B15"].Formula = "PIVOTBY('FCL V'!C6:C2055,'FCL V'!Y6:Y2055,'FCL V'!DH6:DH2055, _xleta.SUM)";
                sheet.Calculate();

                Assert.AreEqual("Albania", sheet.Cells["C15"].Value);
                SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void PivotByTemplateTest2()
        {
            using (var package = OpenTemplatePackage("PivotByTest1.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[3];
                package.Workbook.CalcMode = ExcelCalcMode.Manual;

                sheet.Cells["B26"].Formula = "PIVOTBY('FCL V'!C6:D2055,'FCL V'!Y6:Y2055,'FCL V'!DH6:DH2055, _xleta.PERCENTOF,,2,,,,,3)";
                sheet.Cells["B26"].Calculate();
                
                Assert.AreEqual("Albania", sheet.Cells["D26"].Value);
                Assert.AreEqual(0.021928991, System.Math.Round((double)sheet.Cells["G27"].Value), 8);

                SaveAndCleanup(package);
            }
        }


        [TestMethod]
        public void PivotBySortOrderPercentOf()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["A4"].Value = "B";
                s.Cells["A5"].Value = "C";
                s.Cells["A6"].Value = "C";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["B4"].Value = "Y";
                s.Cells["B5"].Value = "X";
                s.Cells["B6"].Value = "Y";
                s.Cells["C1"].Value = 2;
                s.Cells["C2"].Value = 4;
                s.Cells["C3"].Value = 1;
                s.Cells["C4"].Value = 5;
                s.Cells["C5"].Value = 7;
                s.Cells["C6"].Value = 4;
                s.Cells["D1"].Formula = "PIVOTBY(A1:A6,B1:B6,C1:C6,_xleta.PERCENTOF)";
                s.Calculate();
                Assert.AreEqual(0.2, s.Cells["E2"].Value);
            }
        }

        [TestMethod]
        public void PivotBy_ParentColTotal_ThreeColumnLevels_UsesFullParentPrefix()
        {
            // Verifies that PERCENTOF with RelativeTo=ParentColTotal (3) uses the full
            // parent path (Year, Quarter) as the denominator group when there are three
            // column levels, not just the top-level Year.
            //
            // Current implementation in ResolveRelativeToValues builds parentKey from
            // Path[0] only:
            //     var parentKey = colLeaf.Path[0]?.ToString()?.ToLowerInvariant() ?? "";
            // For three levels this groups by Year alone, producing 10/60 = 0.1667 for
            // the (R1, 2025/Q1/Jan) cell. Excel returns 10/(10+20) = 0.3333, grouping
            // by the full parent prefix (Year, Quarter).
            //
            // Data: R1 has Jan=10 and Feb=20 under 2025/Q1, and Apr=30 under 2025/Q2.
            // R2 has Jan=40 under 2026/Q1.
            //
            // Verified in Excel (sv-SE) 2026-05-21:
            //   Spill range:   G1:L6
            //   Column header layout (H..L):
            //     Year:    2025 2025 2025 2026 Total
            //     Quarter: Q1   Q1   Q2   Q1   -
            //     Month:   Feb  Jan  Apr  Jan  -
            //   R1 row (row 4): G4="R1", H4=0.666667, I4=0.333333, J4=1, K4=<blank>, L4=1
            //
            // Expected to FAIL against current EPPlus implementation.
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "R1";
                s.Cells["A2"].Value = "R1";
                s.Cells["A3"].Value = "R1";
                s.Cells["A4"].Value = "R2";
                s.Cells["B1"].Value = 2025;
                s.Cells["B2"].Value = 2025;
                s.Cells["B3"].Value = 2025;
                s.Cells["B4"].Value = 2026;
                s.Cells["C1"].Value = "Q1";
                s.Cells["C2"].Value = "Q1";
                s.Cells["C3"].Value = "Q2";
                s.Cells["C4"].Value = "Q1";
                s.Cells["D1"].Value = "Jan";
                s.Cells["D2"].Value = "Feb";
                s.Cells["D3"].Value = "Apr";
                s.Cells["D4"].Value = "Jan";
                s.Cells["E1"].Value = 10;
                s.Cells["E2"].Value = 20;
                s.Cells["E3"].Value = 30;
                s.Cells["E4"].Value = 40;

                s.Cells["G1"].Formula = "PIVOTBY(A1:A4, B1:D4, E1:E4, _xleta.PERCENTOF,,,,,,,3)";
                s.Calculate();

                // --- Column header layout (rows 1-3) ---
                Assert.AreEqual(2025, s.Cells["H1"].Value, "H1 Year");
                Assert.AreEqual(2025, s.Cells["I1"].Value, "I1 Year");
                Assert.AreEqual(2025, s.Cells["J1"].Value, "J1 Year");
                Assert.AreEqual(2026, s.Cells["K1"].Value, "K1 Year");
                Assert.AreEqual("Total", s.Cells["L1"].Value, "L1 Total label");

                Assert.AreEqual("Q1", s.Cells["H2"].Value, "H2 Quarter");
                Assert.AreEqual("Q1", s.Cells["I2"].Value, "I2 Quarter");
                Assert.AreEqual("Q2", s.Cells["J2"].Value, "J2 Quarter");
                Assert.AreEqual("Q1", s.Cells["K2"].Value, "K2 Quarter");

                Assert.AreEqual("Feb", s.Cells["H3"].Value, "H3 Month");
                Assert.AreEqual("Jan", s.Cells["I3"].Value, "I3 Month");
                Assert.AreEqual("Apr", s.Cells["J3"].Value, "J3 Month");
                Assert.AreEqual("Jan", s.Cells["K3"].Value, "K3 Month");

                // --- R1 row (row 4) ---
                Assert.AreEqual("R1", s.Cells["G4"].Value, "Row label for R1 should be at G4.");

                // Feb under (2025, Q1): 20 / (10 + 20) = 0.6667
                Assert.AreEqual(0.666667, System.Math.Round((double)s.Cells["H4"].Value, 6),
                    "Feb under (2025,Q1) - denominator must be parent group (Year,Quarter), not Year alone.");

                // Jan under (2025, Q1): 10 / (10 + 20) = 0.3333
                Assert.AreEqual(0.333333, System.Math.Round((double)s.Cells["I4"].Value, 6),
                    "Jan under (2025,Q1) - denominator must be parent group (Year,Quarter), not Year alone.");

                // Apr under (2025, Q2): 30 / 30 = 1.0 (sole value in its parent group)
                Assert.AreEqual(1d, (double)s.Cells["J4"].Value,
                    "Apr under (2025,Q2) is the only value in its parent group, so percent = 1.");

                // K4: R1 has no value under (2026, Q1, Jan) - cell should be blank.
                Assert.IsTrue(
                    s.Cells["K4"].Value == null || (s.Cells["K4"].Value as string) == string.Empty,
                    "K4 should be blank since R1 has no value under (2026,Q1,Jan). Got: " + (s.Cells["K4"].Value ?? "null"));

                // L4: Row total for R1 = 60/60 = 1.
                Assert.AreEqual(1d, (double)s.Cells["L4"].Value,
                    "Row total for R1 should be 1 (whole / whole).");
            }
        }



        [TestMethod]
        public void PivotBy_FilterArray_NumericValuesProducedByExpression()
        {
            // BuildPivotData filters rows with:
            //     if (fv is bool b && !b) continue;
            //     if (fv is int i && i == 0) continue;
            // Excel ranges typically yield doubles for numeric cells, and an expression
            // like (range > x) * 1 produces double 0.0 / 1.0 - matching neither check.
            // The suspicion is that EPPlus silently fails to exclude rows here.
            //
            // Verified in Excel (sv-SE) 2026-05-21: this formula behaves IDENTICALLY
            // to the boolean version above - row A is excluded, grand total = 50.
            //   Spill range: J1:L4
            //     Row 1:        ""    "X"   "Total"
            //     Row 2: "B"    20          20
            //     Row 3: "C"    30          30
            //     Row 4: "Total" 50         50
            //
            // Expected to FAIL against current EPPlus implementation if the bool/int
            // check is the only filter path. If it passes, the bug doesn't exist and
            // the test serves as a regression guard.
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "C";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "X";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = 10;
                s.Cells["C2"].Value = 20;
                s.Cells["C3"].Value = 30;
                // Multiplying by 1 coerces booleans to numerics: {0; 1; 1} as doubles.
                s.Cells["J1"].Formula = "PIVOTBY(A1:A3, B1:B3, C1:C3, _xleta.SUM,,,,,, (C1:C3>15)*1)";
                s.Calculate();

                Assert.AreEqual("X", s.Cells["K1"].Value);
                Assert.AreEqual("Total", s.Cells["L1"].Value);

                Assert.AreEqual("B", s.Cells["J2"].Value);
                Assert.AreEqual(20d, s.Cells["K2"].Value);
                Assert.AreEqual(20d, s.Cells["L2"].Value);

                Assert.AreEqual("C", s.Cells["J3"].Value);
                Assert.AreEqual(30d, s.Cells["K3"].Value);
                Assert.AreEqual(30d, s.Cells["L3"].Value);

                Assert.AreEqual("Total", s.Cells["J4"].Value,
                    "Numeric filter array should also exclude row A - grand total should be 50, not 60.");
                Assert.AreEqual(50d, s.Cells["K4"].Value,
                    "X column total should be 50 (B+C), not 60 (A+B+C).");
                Assert.AreEqual(50d, s.Cells["L4"].Value,
                    "Grand total should be 50, not 60.");
            }
        }

        [TestMethod]
        public void PivotBy_NegativeRowTotalDepth_GrandTotalAtTop()
        {
            // RowTotalDepth = -1 puts the row grand total ABOVE the data rows.
            // This exercises the rowTotalAtTop branch in RenderPivot, which is
            // untested in the existing test suite.
            //
            // Verified in Excel (sv-SE) 2026-05-21:
            //   Spill range: E1:G5
            //     Row 1:        ""     "X"   "Total"
            //     Row 2: "Total" 6     6
            //     Row 3: "A"     1     1
            //     Row 4: "B"     2     2
            //     Row 5: "C"     3     3
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "C";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "X";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = 1;
                s.Cells["C2"].Value = 2;
                s.Cells["C3"].Value = 3;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3, B1:B3, C1:C3, _xleta.SUM,,-1)";
                s.Calculate();

                // --- Header row ---
                Assert.AreEqual("X", s.Cells["F1"].Value);
                Assert.AreEqual("Total", s.Cells["G1"].Value);

                // --- Grand total at TOP (row 2) ---
                Assert.AreEqual("Total", s.Cells["E2"].Value,
                    "With RowTotalDepth=-1 the grand total row must appear above the data rows.");
                Assert.AreEqual(6d, s.Cells["F2"].Value);
                Assert.AreEqual(6d, s.Cells["G2"].Value);

                // --- Data rows (3-5) ---
                Assert.AreEqual("A", s.Cells["E3"].Value);
                Assert.AreEqual(1d, s.Cells["F3"].Value);
                Assert.AreEqual(1d, s.Cells["G3"].Value);

                Assert.AreEqual("B", s.Cells["E4"].Value);
                Assert.AreEqual(2d, s.Cells["F4"].Value);
                Assert.AreEqual(2d, s.Cells["G4"].Value);

                Assert.AreEqual("C", s.Cells["E5"].Value);
                Assert.AreEqual(3d, s.Cells["F5"].Value);
                Assert.AreEqual(3d, s.Cells["G5"].Value);
            }
        }

        [TestMethod]
        public void PivotBy_NegativeColTotalDepth_GrandTotalAtLeft()
        {
            // ColTotalDepth = -1 puts the column grand total LEFT of the data columns.
            // This exercises the colTotalAtLeft branch in RenderPivot, which is
            // untested in the existing test suite.
            //
            // Verified in Excel (sv-SE) 2026-05-21:
            //   Spill range: E1:H5
            //     Row 1:        ""      "Total"  "X"   "Y"
            //     Row 2: "A"     1      1
            //     Row 3: "B"     2              2
            //     Row 4: "C"     3      3
            //     Row 5: "Total" 6      4       2
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "C";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["C1"].Value = 1;
                s.Cells["C2"].Value = 2;
                s.Cells["C3"].Value = 3;
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3, B1:B3, C1:C3, _xleta.SUM,,,,-1)";
                s.Calculate();

                // --- Header row: Total column comes FIRST (left of X and Y) ---
                Assert.AreEqual("Total", s.Cells["F1"].Value,
                    "With ColTotalDepth=-1 the column total must appear leftmost.");
                Assert.AreEqual("X", s.Cells["G1"].Value);
                Assert.AreEqual("Y", s.Cells["H1"].Value);

                // --- Row A (only has X) ---
                Assert.AreEqual("A", s.Cells["E2"].Value);
                Assert.AreEqual(1d, s.Cells["F2"].Value, "Row A total at leftmost column.");
                Assert.AreEqual(1d, s.Cells["G2"].Value, "A under X.");
                Assert.IsTrue(
                    s.Cells["H2"].Value == null || (s.Cells["H2"].Value as string) == string.Empty,
                    "A has no Y value - H2 should be blank.");

                // --- Row B (only has Y) ---
                Assert.AreEqual("B", s.Cells["E3"].Value);
                Assert.AreEqual(2d, s.Cells["F3"].Value, "Row B total at leftmost column.");
                Assert.IsTrue(
                    s.Cells["G3"].Value == null || (s.Cells["G3"].Value as string) == string.Empty,
                    "B has no X value - G3 should be blank.");
                Assert.AreEqual(2d, s.Cells["H3"].Value, "B under Y.");

                // --- Row C (only has X) ---
                Assert.AreEqual("C", s.Cells["E4"].Value);
                Assert.AreEqual(3d, s.Cells["F4"].Value, "Row C total at leftmost column.");
                Assert.AreEqual(3d, s.Cells["G4"].Value, "C under X.");

                // --- Grand total row ---
                Assert.AreEqual("Total", s.Cells["E5"].Value);
                Assert.AreEqual(6d, s.Cells["F5"].Value, "Grand total of grand totals.");
                Assert.AreEqual(4d, s.Cells["G5"].Value, "X column total = A+C.");
                Assert.AreEqual(2d, s.Cells["H5"].Value, "Y column total = B.");
            }
        }

        [TestMethod]
        public void PivotBy_WithoutFunctionArgument_ShouldReturnErrorNotThrow()
        {
            // ArgumentMinLength is declared as 3 in PivotBy, but TryParsePivotByArgs
            // accesses arguments[3] unconditionally:
            //     if (!TryParseFunctionArg(arguments[3], ...))
            //
            // With only 3 arguments this throws IndexOutOfRangeException instead of
            // returning a proper error value. Excel itself rejects the formula at
            // parse time with "Too few arguments", confirming that 4 is the real
            // minimum. EPPlus should either bump ArgumentMinLength to 4 or guard
            // the access and return #VALUE!.
            //
            // Verified in Excel (sv-SE) 2026-05-21: formula rejected at parse time
            // ("Too few arguments").
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["C1"].Value = 1;
                s.Cells["C2"].Value = 2;
                s.Cells["D1"].Formula = "PIVOTBY(A1:A2, B1:B2, C1:C2)";

                // Must not throw - the calculation should complete and produce an
                // error value in the cell.
                try
                {
                    s.Calculate();
                }
                catch (Exception ex)
                {
                    Assert.Fail(
                        "PIVOTBY with too few arguments must not throw - it should " +
                        "produce an error value. Got: " + ex.GetType().Name +
                        ": " + ex.Message);
                }

                var value = s.Cells["D1"].Value;
                Assert.IsInstanceOfType(
                    value,
                    typeof(ExcelErrorValue),
                    "Expected an ExcelErrorValue (e.g. #VALUE!) when function argument " +
                    "is omitted, got: " + (value == null ? "null" : value.GetType().Name + " = " + value));
            }
        }
    }
}
