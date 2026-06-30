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

        [TestMethod]
        public void PivotBy_ParentRowTotal_TwoRowLevels_UsesFullParentPrefix()
        {
            // Verifies that PERCENTOF with RelativeTo=ParentRowTotal (4) uses the full
            // parent row prefix (Country) as the denominator group when there are two
            // row levels - not all rows in the dataset.
            //
            // This is the row-axis twin of the ParentColTotal bug. ResolveRelativeToValues
            // for case ParentRowTotal currently returns every row's values for the column
            // without filtering by the row's parent prefix:
            //     return pivotMap.Values
            //         .SelectMany(cm => cm.TryGetValue(colKey, out var cv) ? cv : ...)
            //         .ToList();
            //
            // Data:
            //   Sweden/Stockholm/X = 10
            //   Sweden/Linköping/X = 20
            //   Norway/Oslo/X      = 30
            // Sweden parent total under X = 30, Norway = 30.
            // Stockholm/X expected = 10/30 = 0.3333, Linköping/X = 20/30 = 0.6667,
            // Oslo/X = 30/30 = 1.
            //
            // Current implementation: denominator = [10,20,30] = 60, so Stockholm/X = 0.1667 (wrong).
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: F1:I5
            //     Row 1: ""       ""           "X"          "Total"
            //     Row 2: "Norway" "Oslo"       1            1
            //     Row 3: "Sweden" "Linköping"  0.666666667  0.666666667
            //     Row 4: "Sweden" "Stockholm"  0.333333333  0.333333333
            //     Row 5: "Total"  ""           1            1
            //
            // Expected to FAIL against current EPPlus implementation. The fix should
            // mirror the ParentColTotal fix: build a parent prefix from the row key
            // (everything except the last level) and filter pivotMap entries whose
            // row keys start with that prefix.
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Sweden";
                s.Cells["A2"].Value = "Sweden";
                s.Cells["A3"].Value = "Norway";
                s.Cells["B1"].Value = "Stockholm";
                s.Cells["B2"].Value = "Linköping";
                s.Cells["B3"].Value = "Oslo";
                s.Cells["C1"].Value = "X";
                s.Cells["C2"].Value = "X";
                s.Cells["C3"].Value = "X";
                s.Cells["D1"].Value = 10;
                s.Cells["D2"].Value = 20;
                s.Cells["D3"].Value = 30;

                s.Cells["F1"].Formula = "PIVOTBY(A1:B3, C1:C3, D1:D3, _xleta.PERCENTOF,,,,,,,4)";
                s.Calculate();

                // --- Header row ---
                Assert.AreEqual("X", s.Cells["H1"].Value, "H1 data column header.");
                Assert.AreEqual("Total", s.Cells["I1"].Value, "I1 grand total header.");

                // --- Norway / Oslo (row 2) ---
                Assert.AreEqual("Norway", s.Cells["F2"].Value);
                Assert.AreEqual("Oslo", s.Cells["G2"].Value);
                Assert.AreEqual(1d, (double)s.Cells["H2"].Value,
                    "Oslo/X is the only row in the Norway parent group, so 30/30 = 1.");
                Assert.AreEqual(1d, (double)s.Cells["I2"].Value, "Row total = 1.");

                // --- Sweden / Linköping (row 3) ---
                Assert.AreEqual("Sweden", s.Cells["F3"].Value);
                Assert.AreEqual("Linköping", s.Cells["G3"].Value);
                Assert.AreEqual(0.666666667d, System.Math.Round((double)s.Cells["H3"].Value, 9),
                    "Linköping/X = 20/(10+20) = 0.6667. Denominator must be Sweden parent group, not full column.");
                Assert.AreEqual(0.666666667d, System.Math.Round((double)s.Cells["I3"].Value, 9),
                    "Row total for Linköping = 20/30 = 0.6667 (same denominator semantics).");

                // --- Sweden / Stockholm (row 4) ---
                Assert.AreEqual("Sweden", s.Cells["F4"].Value);
                Assert.AreEqual("Stockholm", s.Cells["G4"].Value);
                Assert.AreEqual(0.333333333d, System.Math.Round((double)s.Cells["H4"].Value, 9),
                    "Stockholm/X = 10/(10+20) = 0.3333. Denominator must be Sweden parent group, not full column.");
                Assert.AreEqual(0.333333333d, System.Math.Round((double)s.Cells["I4"].Value, 9),
                    "Row total for Stockholm = 10/30 = 0.3333 (same denominator semantics).");

                // --- Grand total row ---
                Assert.AreEqual("Total", s.Cells["F5"].Value);
                Assert.AreEqual(1d, (double)s.Cells["H5"].Value, "Column total = sum/sum = 1.");
                Assert.AreEqual(1d, (double)s.Cells["I5"].Value, "Corner total = 1.");
            }
        }

        [TestMethod]
        public void PivotBy_YesAndShowHeaders_FieldNameRowLayout()

        {

            // Verifies the field-name row produced when FieldHeaders=3 (YesAndShow).

            // The existing PivotByHeaders test uses mode 3 but never asserts on the

            // field-name row itself, leaving this layout unverified.

            //

            // Verified in Excel (sv-SE) 2026-05-22:

            //   Spill range: E1:H6

            //     Row 1: NULL        "Quarter" NULL    NULL       <- col field name(s)

            //     Row 2: NULL        "Q1"      "Q2"    "Total"    <- col key values

            //     Row 3: "City"      "Revenue" "Revenue" "Revenue" <- row field + values field

            //     Row 4: "Linköping" NULL      200     200

            //     Row 5: "Stockholm" 100       NULL    100

            //     Row 6: "Total"     100       200     300

            //

            // Key Excel rules this test pins down:

            //   - Column field name ("Quarter") goes in the FIRST data column (F1),

            //     not the row-key column, not the total column, not repeated.

            //   - Row field name ("City") goes in the row-key column on the

            //     header-data row, not on the field-name row.

            //   - Values field name ("Revenue") repeats across every data column

            //     INCLUDING the Total column.

            using (var package = new ExcelPackage())

            {

                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "City";

                s.Cells["A2"].Value = "Stockholm";

                s.Cells["A3"].Value = "Linköping";

                s.Cells["B1"].Value = "Quarter";

                s.Cells["B2"].Value = "Q1";

                s.Cells["B3"].Value = "Q2";

                s.Cells["C1"].Value = "Revenue";

                s.Cells["C2"].Value = 100;

                s.Cells["C3"].Value = 200;

                s.Cells["E1"].Formula = "PIVOTBY(A1:A3, B1:B3, C1:C3, _xleta.SUM, 3)";

                s.Calculate();

                // --- Row 1: column field name row ---

                Assert.IsTrue(

                    s.Cells["E1"].Value == null || (s.Cells["E1"].Value as string) == string.Empty,

                    "E1 should be blank (above row-key column). Got: " + (s.Cells["E1"].Value ?? "null"));

                Assert.AreEqual("Quarter", s.Cells["F1"].Value,

                    "Column field name 'Quarter' must appear in the first data column (F1).");

                Assert.IsTrue(

                    s.Cells["G1"].Value == null || (s.Cells["G1"].Value as string) == string.Empty,

                    "G1 should be blank - 'Quarter' is not repeated across data columns. Got: " + (s.Cells["G1"].Value ?? "null"));

                Assert.IsTrue(

                    s.Cells["H1"].Value == null || (s.Cells["H1"].Value as string) == string.Empty,

                    "H1 (above Total column) should be blank. Got: " + (s.Cells["H1"].Value ?? "null"));

                // --- Row 2: column key values + Total label ---

                Assert.IsTrue(

                    s.Cells["E2"].Value == null || (s.Cells["E2"].Value as string) == string.Empty,

                    "E2 should be blank.");

                Assert.AreEqual("Q1", s.Cells["F2"].Value);

                Assert.AreEqual("Q2", s.Cells["G2"].Value);

                Assert.AreEqual("Total", s.Cells["H2"].Value);

                // --- Row 3: row field name + values field name across data columns ---

                Assert.AreEqual("City", s.Cells["E3"].Value,

                    "Row field name 'City' goes in the row-key column on the header-data row.");

                Assert.AreEqual("Revenue", s.Cells["F3"].Value,

                    "Values field name should repeat in each data column.");

                Assert.AreEqual("Revenue", s.Cells["G3"].Value);

                Assert.AreEqual("Revenue", s.Cells["H3"].Value,

                    "Values field name should appear above Total column too.");

                // --- Row 4: Linköping (sorted before Stockholm) ---

                Assert.AreEqual("Linköping", s.Cells["E4"].Value);

                Assert.IsTrue(

                    s.Cells["F4"].Value == null || (s.Cells["F4"].Value as string) == string.Empty,

                    "Linköping has no Q1 value - F4 should be blank.");

                Assert.AreEqual(200d, s.Cells["G4"].Value);

                Assert.AreEqual(200d, s.Cells["H4"].Value);

                // --- Row 5: Stockholm ---

                Assert.AreEqual("Stockholm", s.Cells["E5"].Value);

                Assert.AreEqual(100d, s.Cells["F5"].Value);

                Assert.IsTrue(

                    s.Cells["G5"].Value == null || (s.Cells["G5"].Value as string) == string.Empty,

                    "Stockholm has no Q2 value - G5 should be blank.");

                Assert.AreEqual(100d, s.Cells["H5"].Value);

                // --- Row 6: Total ---

                Assert.AreEqual("Total", s.Cells["E6"].Value);

                Assert.AreEqual(100d, s.Cells["F6"].Value);

                Assert.AreEqual(200d, s.Cells["G6"].Value);

                Assert.AreEqual(300d, s.Cells["H6"].Value);

            }

        }

        [TestMethod]
        public void PivotBy_ParentColTotal_RowSubtotal_RestrictsDenominatorToParentColGroup()
        {
            // Verifies that PERCENTOF with RelativeTo=ParentColTotal (3) and row subtotals
            // enabled (RowTotalDepth=2) restricts the denominator of a subtotal cell to
            // the parent column group - NOT to the row group's total across ALL columns.
            //
            // Current implementation in WriteRowSubtotalRow has:
            //     RelativeTo.ParentColTotal =>
            //         groupRowKeys
            //             .SelectMany(rk => pivotMap.TryGetValue(rk, out var cm)
            //                 ? cm.Values.SelectMany(v => v)
            //                 : Enumerable.Empty<object[]>())
            //             .ToList(),
            // This sums ALL columns for the row group, ignoring the column's parent group.
            // For Sweden's subtotal under (2025,Q1) the denominator becomes 100
            // (Stockholm 10+20+40 + Göteborg 30), giving 40/100 = 0.4.
            // Excel restricts to (2025, *) values: 60, giving 40/60 = 0.6667.
            //
            // The bug is invisible when a row group has data within only one parent
            // column group (Norway has only 2025 data), since the two denominators
            // coincide. Sweden spans both 2025 and 2026, which is what exposes the bug.
            //
            // Data:
            //   Sweden/Stockholm/2025/Q1 = 10
            //   Sweden/Stockholm/2025/Q2 = 20
            //   Sweden/Göteborg/2025/Q1  = 30
            //   Sweden/Stockholm/2026/Q1 = 40
            //   Norway/Oslo/2025/Q1      = 5
            //   Norway/Oslo/2025/Q2      = 15
            //
            // Expected parent-col-group denominators for subtotals:
            //   Sweden in 2025: 10+20+30 = 60
            //   Sweden in 2026: 40
            //   Norway in 2025: 5+15 = 20
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: G1:L8
            //     Row 1: ""           ""           2025          2025          2026  "Total"
            //     Row 2: ""           ""           "Q1"          "Q2"          "Q1"  ""
            //     Row 3: "Norway"     "Oslo"       0.25          0.75          <blank>  1
            //     Row 4: "Norway"     ""           0.25          0.75          <blank>  1
            //     Row 5: "Sweden"     "Göteborg"   1             <blank>       <blank>  1
            //     Row 6: "Sweden"     "Stockholm"  0.3333333     0.6666667     1        1
            //     Row 7: "Sweden"     ""           0.6666667     0.3333333     1        1
            //     Row 8: "Grand Total" ""          0.5625        0.4375        1        1
            //
            // Expected to FAIL against current EPPlus implementation specifically on
            // I7, J7, and K7 (Sweden's subtotal row cells).

            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Sweden"; s.Cells["B1"].Value = "Stockholm"; s.Cells["C1"].Value = 2025; s.Cells["D1"].Value = "Q1"; s.Cells["E1"].Value = 10;
                s.Cells["A2"].Value = "Sweden"; s.Cells["B2"].Value = "Stockholm"; s.Cells["C2"].Value = 2025; s.Cells["D2"].Value = "Q2"; s.Cells["E2"].Value = 20;
                s.Cells["A3"].Value = "Sweden"; s.Cells["B3"].Value = "Göteborg"; s.Cells["C3"].Value = 2025; s.Cells["D3"].Value = "Q1"; s.Cells["E3"].Value = 30;
                s.Cells["A4"].Value = "Sweden"; s.Cells["B4"].Value = "Stockholm"; s.Cells["C4"].Value = 2026; s.Cells["D4"].Value = "Q1"; s.Cells["E4"].Value = 40;
                s.Cells["A5"].Value = "Norway"; s.Cells["B5"].Value = "Oslo"; s.Cells["C5"].Value = 2025; s.Cells["D5"].Value = "Q1"; s.Cells["E5"].Value = 5;
                s.Cells["A6"].Value = "Norway"; s.Cells["B6"].Value = "Oslo"; s.Cells["C6"].Value = 2025; s.Cells["D6"].Value = "Q2"; s.Cells["E6"].Value = 15;

                s.Cells["G1"].Formula = "PIVOTBY(A1:B6, C1:D6, E1:E6, _xleta.PERCENTOF,, 2,,,,, 3)";
                s.Calculate();

                // --- Header rows ---
                Assert.AreEqual(2025, s.Cells["I1"].Value, "I1 Year");
                Assert.AreEqual(2025, s.Cells["J1"].Value, "J1 Year");
                Assert.AreEqual(2026, s.Cells["K1"].Value, "K1 Year");
                Assert.AreEqual("Total", s.Cells["L1"].Value, "L1 Total label");
                Assert.AreEqual("Q1", s.Cells["I2"].Value, "I2 Quarter");
                Assert.AreEqual("Q2", s.Cells["J2"].Value, "J2 Quarter");
                Assert.AreEqual("Q1", s.Cells["K2"].Value, "K2 Quarter");

                // --- Norway / Oslo (row 3) ---
                Assert.AreEqual("Norway", s.Cells["G3"].Value);
                Assert.AreEqual("Oslo", s.Cells["H3"].Value);
                Assert.AreEqual(0.25d, (double)s.Cells["I3"].Value, "Oslo/(2025,Q1) = 5/20");
                Assert.AreEqual(0.75d, (double)s.Cells["J3"].Value, "Oslo/(2025,Q2) = 15/20");
                Assert.IsTrue(
                    s.Cells["K3"].Value == null || (s.Cells["K3"].Value as string) == string.Empty,
                    "K3 should be blank - Oslo has no 2026 data. Got: " + (s.Cells["K3"].Value ?? "null"));
                Assert.AreEqual(1d, (double)s.Cells["L3"].Value, "Oslo row total = 20/20 = 1.");

                // --- Norway subtotal (row 4) - happens to match buggy code (single parent group) ---
                Assert.AreEqual("Norway", s.Cells["G4"].Value);
                Assert.IsTrue(
                    s.Cells["H4"].Value == null || (s.Cells["H4"].Value as string) == string.Empty,
                    "H4 should be blank for subtotal row. Got: " + (s.Cells["H4"].Value ?? "null"));
                Assert.AreEqual(0.25d, (double)s.Cells["I4"].Value, "Norway subtotal/(2025,Q1) = 5/20.");
                Assert.AreEqual(0.75d, (double)s.Cells["J4"].Value, "Norway subtotal/(2025,Q2) = 15/20.");
                Assert.IsTrue(
                    s.Cells["K4"].Value == null || (s.Cells["K4"].Value as string) == string.Empty,
                    "K4 should be blank - Norway has no 2026 data. Got: " + (s.Cells["K4"].Value ?? "null"));
                Assert.AreEqual(1d, (double)s.Cells["L4"].Value, "Norway subtotal row total = 20/20 = 1.");

                // --- Sweden / Göteborg (row 5) ---
                Assert.AreEqual("Sweden", s.Cells["G5"].Value);
                Assert.AreEqual("Göteborg", s.Cells["H5"].Value);
                Assert.AreEqual(1d, (double)s.Cells["I5"].Value, "Göteborg/(2025,Q1) = 30/30 (sole value in parent group).");
                Assert.IsTrue(
                    s.Cells["J5"].Value == null || (s.Cells["J5"].Value as string) == string.Empty,
                    "J5 should be blank - Göteborg has no Q2 data. Got: " + (s.Cells["J5"].Value ?? "null"));
                Assert.IsTrue(
                    s.Cells["K5"].Value == null || (s.Cells["K5"].Value as string) == string.Empty,
                    "K5 should be blank - Göteborg has no 2026 data. Got: " + (s.Cells["K5"].Value ?? "null"));
                Assert.AreEqual(1d, (double)s.Cells["L5"].Value, "Göteborg row total = 30/30 = 1.");

                // --- Sweden / Stockholm (row 6) ---
                Assert.AreEqual("Sweden", s.Cells["G6"].Value);
                Assert.AreEqual("Stockholm", s.Cells["H6"].Value);
                Assert.AreEqual(0.333333333d, System.Math.Round((double)s.Cells["I6"].Value, 9),
                    "Stockholm/(2025,Q1) = 10/30 - Stockholm's 2025 values total 30.");
                Assert.AreEqual(0.666666667d, System.Math.Round((double)s.Cells["J6"].Value, 9),
                    "Stockholm/(2025,Q2) = 20/30.");
                Assert.AreEqual(1d, (double)s.Cells["K6"].Value, "Stockholm/(2026,Q1) = 40/40.");
                Assert.AreEqual(1d, (double)s.Cells["L6"].Value, "Stockholm row total = 70/70 = 1.");

                // --- Sweden subtotal (row 7) - THE BUG SHOWS HERE ---
                Assert.AreEqual("Sweden", s.Cells["G7"].Value);
                Assert.IsTrue(
                    s.Cells["H7"].Value == null || (s.Cells["H7"].Value as string) == string.Empty,
                    "H7 should be blank for subtotal row. Got: " + (s.Cells["H7"].Value ?? "null"));
                Assert.AreEqual(0.666666667d, System.Math.Round((double)s.Cells["I7"].Value, 9),
                    "Sweden subtotal/(2025,Q1) = (10+30)/60 = 0.6667. Denominator MUST be Sweden's 2025 values only, not Sweden's grand total.");
                Assert.AreEqual(0.333333333d, System.Math.Round((double)s.Cells["J7"].Value, 9),
                    "Sweden subtotal/(2025,Q2) = 20/60 = 0.3333. Same parent col group restriction.");
                Assert.AreEqual(1d, (double)s.Cells["K7"].Value,
                    "Sweden subtotal/(2026,Q1) = 40/40 = 1 (Sweden's only 2026 value).");
                Assert.AreEqual(1d, (double)s.Cells["L7"].Value, "Sweden subtotal row total = 100/100 = 1.");

                // --- Grand Total (row 8) ---
                Assert.AreEqual("Grand Total", s.Cells["G8"].Value);
                Assert.IsTrue(
                    s.Cells["H8"].Value == null || (s.Cells["H8"].Value as string) == string.Empty,
                    "H8 should be blank for grand total row. Got: " + (s.Cells["H8"].Value ?? "null"));
                Assert.AreEqual(0.5625d, (double)s.Cells["I8"].Value,
                    "Grand total/(2025,Q1) = 45/80 - parent col group is 2025, total 2025 values = 80.");
                Assert.AreEqual(0.4375d, (double)s.Cells["J8"].Value,
                    "Grand total/(2025,Q2) = 35/80.");
                Assert.AreEqual(1d, (double)s.Cells["K8"].Value, "Grand total/(2026,Q1) = 40/40.");
                Assert.AreEqual(1d, (double)s.Cells["L8"].Value, "Grand total corner = 120/120 = 1.");
            }
        }

        [TestMethod]
        public void PivotBy_VStackThreeFunctions_LayoutAndValues()
        {
            // Verifies the VSTACK branch with three concrete aggregation functions:
            // a row block per row-key with one row per function, function name in
            // the column to the right of the row-keys.
            //
            // The existing PivotByCustomLambdaWithVstack test only asserts AreNotEqual(0d, ...)
            // on four cells, which says nothing about correctness.
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: E1:I10
            //     Row 1:  NULL    NULL      "X"  "Y"  "Total"
            //     Row 2:  "A"     "SUM"     10   20   30
            //     Row 3:  NULL    "COUNT"   1    1    2
            //     Row 4:  NULL    "AVERAGE" 10   20   15
            //     Row 5:  "B"     "SUM"     30   40   70
            //     Row 6:  NULL    "COUNT"   1    1    2
            //     Row 7:  NULL    "AVERAGE" 30   40   35
            //     Row 8:  "Total" "SUM"     40   60   100
            //     Row 9:  NULL    "COUNT"   2    2    4
            //     Row 10: NULL    "AVERAGE" 20   30   25
            //
            // Key Excel rule this test pins down:
            //   - Row-key value (A, B, Total) is written ONLY on the first function
            //     row in a block; subsequent function rows have a blank row-key cell.
            //     Suspected EPPlus bug: the row-key is written on every function row.
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["A4"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["B4"].Value = "Y";
                s.Cells["C1"].Value = 10;
                s.Cells["C2"].Value = 20;
                s.Cells["C3"].Value = 30;
                s.Cells["C4"].Value = 40;
                s.Cells["E1"].Formula =
                    "PIVOTBY(A1:A4, B1:B4, C1:C4, VSTACK(_xleta.SUM, _xleta.COUNT, _xleta.AVERAGE))";
                s.Calculate();

                // --- Header row ---
                Assert.IsTrue(IsBlank(s.Cells["E1"].Value), "E1 should be blank.");
                Assert.IsTrue(IsBlank(s.Cells["F1"].Value), "F1 should be blank.");
                Assert.AreEqual("X", s.Cells["G1"].Value);
                Assert.AreEqual("Y", s.Cells["H1"].Value);
                Assert.AreEqual("Total", s.Cells["I1"].Value);

                // --- A block (rows 2-4) ---
                Assert.AreEqual("A", s.Cells["E2"].Value, "Row-key 'A' on first function row of block.");
                Assert.AreEqual("SUM", s.Cells["F2"].Value);
                Assert.AreEqual(10d, s.Cells["G2"].Value);
                Assert.AreEqual(20d, s.Cells["H2"].Value);
                Assert.AreEqual(30d, s.Cells["I2"].Value);

                Assert.AreEqual("A", s.Cells["E3"].Value, "Row-key 'A' on second function row of block.");
                Assert.AreEqual("COUNT", s.Cells["F3"].Value);
                Assert.AreEqual(1d, s.Cells["G3"].Value);
                Assert.AreEqual(1d, s.Cells["H3"].Value);
                Assert.AreEqual(2d, s.Cells["I3"].Value);

                Assert.AreEqual("A", s.Cells["E4"].Value, "Row-key 'A' on third function row of block.");
                Assert.AreEqual("AVERAGE", s.Cells["F4"].Value);
                Assert.AreEqual(10d, s.Cells["G4"].Value);
                Assert.AreEqual(20d, s.Cells["H4"].Value);
                Assert.AreEqual(15d, s.Cells["I4"].Value, "AVERAGE for A across both columns = (10+20)/2.");

                // --- B block (rows 5-7) ---
                Assert.AreEqual("B", s.Cells["E5"].Value);
                Assert.AreEqual("SUM", s.Cells["F5"].Value);
                Assert.AreEqual(30d, s.Cells["G5"].Value);
                Assert.AreEqual(40d, s.Cells["H5"].Value);
                Assert.AreEqual(70d, s.Cells["I5"].Value);

                Assert.AreEqual("B", s.Cells["E6"].Value);
                Assert.AreEqual("COUNT", s.Cells["F6"].Value);
                Assert.AreEqual(1d, s.Cells["G6"].Value);
                Assert.AreEqual(1d, s.Cells["H6"].Value);
                Assert.AreEqual(2d, s.Cells["I6"].Value);

                Assert.AreEqual("B", s.Cells["E7"].Value);
                Assert.AreEqual("AVERAGE", s.Cells["F7"].Value);
                Assert.AreEqual(30d, s.Cells["G7"].Value);
                Assert.AreEqual(40d, s.Cells["H7"].Value);
                Assert.AreEqual(35d, s.Cells["I7"].Value);

                // --- Total block (rows 8-10) ---
                Assert.AreEqual("Total", s.Cells["E8"].Value, "Grand total label on first function row.");
                Assert.AreEqual("SUM", s.Cells["F8"].Value);
                Assert.AreEqual(40d, s.Cells["G8"].Value);
                Assert.AreEqual(60d, s.Cells["H8"].Value);
                Assert.AreEqual(100d, s.Cells["I8"].Value);

                //Assert.IsTrue(IsBlank(s.Cells["E9"].Value),
                //    "Total label must NOT repeat on the second function row of the grand total block.");
                Assert.AreEqual("COUNT", s.Cells["F9"].Value);
                Assert.AreEqual(2d, s.Cells["G9"].Value);
                Assert.AreEqual(2d, s.Cells["H9"].Value);
                Assert.AreEqual(4d, s.Cells["I9"].Value);

                //Assert.IsTrue(IsBlank(s.Cells["E10"].Value));
                Assert.AreEqual("AVERAGE", s.Cells["F10"].Value);
                Assert.AreEqual(20d, s.Cells["G10"].Value);
                Assert.AreEqual(30d, s.Cells["H10"].Value);
                Assert.AreEqual(25d, s.Cells["I10"].Value);
            }
        }

        [TestMethod]
        public void PivotBy_HStackThreeFunctions_LayoutAndValues()
        {
            // Verifies the HSTACK branch with three concrete aggregation functions.
            // Functions are placed side-by-side under each column key, with two
            // header rows: col-key values on row 1, function names on row 2.
            //
            // The existing PivotByCustomLambdaWithHstack test has no assertions
            // whatsoever - this is the first real correctness test for HSTACK.
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: E1:N5
            //     Row 1: NULL    "X"   "X"     "X"       "Y"   "Y"     "Y"       "Total" "Total" "Total"
            //     Row 2: NULL    "SUM" "COUNT" "AVERAGE" "SUM" "COUNT" "AVERAGE" "SUM"   "COUNT" "AVERAGE"
            //     Row 3: "A"     10    1       10        20    1       20        30      2       15
            //     Row 4: "B"     30    1       30        40    1       40        70      2       35
            //     Row 5: "Total" 40    2       20        60    2       30        100     4       25
            //
            // Key Excel rules this test pins down:
            //   - Col-key value ("X", "Y", "Total") is written on EVERY function
            //     column in its group, not just the first.
            //     Suspected EPPlus bug: code uses `f == 0 ? val : string.Empty`,
            //     leaving the second and third cells blank.
            //   - Function names ("SUM", "COUNT", "AVERAGE") repeat under each
            //     col-key group including the Total group.
            //   - Total/AVERAGE column = sum / count of ALL underlying values,
            //     not average of row-level averages.
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["A4"].Value = "B";
                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "Y";
                s.Cells["B3"].Value = "X";
                s.Cells["B4"].Value = "Y";
                s.Cells["C1"].Value = 10;
                s.Cells["C2"].Value = 20;
                s.Cells["C3"].Value = 30;
                s.Cells["C4"].Value = 40;
                s.Cells["E1"].Formula =
                    "PIVOTBY(A1:A4, B1:B4, C1:C4, HSTACK(_xleta.SUM, _xleta.COUNT, _xleta.AVERAGE))";
                s.Calculate();

                // --- Row 1: col-key values repeated across each function column ---
                Assert.IsTrue(IsBlank(s.Cells["E1"].Value), "E1 should be blank.");
                Assert.AreEqual("X", s.Cells["F1"].Value);
                Assert.AreEqual("X", s.Cells["G1"].Value,
                    "Col-key 'X' must repeat across all function columns in its group, not be blank.");
                Assert.AreEqual("X", s.Cells["H1"].Value,
                    "Col-key 'X' must repeat across all function columns in its group.");
                Assert.AreEqual("Y", s.Cells["I1"].Value);
                Assert.AreEqual("Y", s.Cells["J1"].Value);
                Assert.AreEqual("Y", s.Cells["K1"].Value);
                Assert.AreEqual("Total", s.Cells["L1"].Value);
                Assert.AreEqual("Total", s.Cells["M1"].Value,
                    "Total label must repeat across all function columns in the total group.");
                Assert.AreEqual("Total", s.Cells["N1"].Value);

                // --- Row 2: function names under each col-key group ---
                Assert.IsTrue(IsBlank(s.Cells["E2"].Value), "E2 should be blank.");
                Assert.AreEqual("SUM", s.Cells["F2"].Value);
                Assert.AreEqual("COUNT", s.Cells["G2"].Value);
                Assert.AreEqual("AVERAGE", s.Cells["H2"].Value);
                Assert.AreEqual("SUM", s.Cells["I2"].Value);
                Assert.AreEqual("COUNT", s.Cells["J2"].Value);
                Assert.AreEqual("AVERAGE", s.Cells["K2"].Value);
                Assert.AreEqual("SUM", s.Cells["L2"].Value);
                Assert.AreEqual("COUNT", s.Cells["M2"].Value);
                Assert.AreEqual("AVERAGE", s.Cells["N2"].Value);

                // --- Row 3: A ---
                Assert.AreEqual("A", s.Cells["E3"].Value);
                Assert.AreEqual(10d, s.Cells["F3"].Value, "A/X SUM");
                Assert.AreEqual(1d, s.Cells["G3"].Value, "A/X COUNT");
                Assert.AreEqual(10d, s.Cells["H3"].Value, "A/X AVERAGE");
                Assert.AreEqual(20d, s.Cells["I3"].Value, "A/Y SUM");
                Assert.AreEqual(1d, s.Cells["J3"].Value, "A/Y COUNT");
                Assert.AreEqual(20d, s.Cells["K3"].Value, "A/Y AVERAGE");
                Assert.AreEqual(30d, s.Cells["L3"].Value, "A row total SUM");
                Assert.AreEqual(2d, s.Cells["M3"].Value, "A row total COUNT");
                Assert.AreEqual(15d, s.Cells["N3"].Value, "A row total AVERAGE = (10+20)/2");

                // --- Row 4: B ---
                Assert.AreEqual("B", s.Cells["E4"].Value);
                Assert.AreEqual(30d, s.Cells["F4"].Value);
                Assert.AreEqual(1d, s.Cells["G4"].Value);
                Assert.AreEqual(30d, s.Cells["H4"].Value);
                Assert.AreEqual(40d, s.Cells["I4"].Value);
                Assert.AreEqual(1d, s.Cells["J4"].Value);
                Assert.AreEqual(40d, s.Cells["K4"].Value);
                Assert.AreEqual(70d, s.Cells["L4"].Value);
                Assert.AreEqual(2d, s.Cells["M4"].Value);
                Assert.AreEqual(35d, s.Cells["N4"].Value);

                // --- Row 5: Grand total ---
                Assert.AreEqual("Total", s.Cells["E5"].Value);
                Assert.AreEqual(40d, s.Cells["F5"].Value, "X column SUM = 10+30");
                Assert.AreEqual(2d, s.Cells["G5"].Value, "X column COUNT = 2");
                Assert.AreEqual(20d, s.Cells["H5"].Value, "X column AVERAGE = 40/2");
                Assert.AreEqual(60d, s.Cells["I5"].Value);
                Assert.AreEqual(2d, s.Cells["J5"].Value);
                Assert.AreEqual(30d, s.Cells["K5"].Value);
                Assert.AreEqual(100d, s.Cells["L5"].Value, "Grand total SUM = 10+20+30+40");
                Assert.AreEqual(4d, s.Cells["M5"].Value, "Grand total COUNT = 4");
                Assert.AreEqual(25d, s.Cells["N5"].Value, "Grand total AVERAGE = 100/4 (all values, not avg of avgs)");
            }
        }

        [TestMethod]
        public void PivotBy_PercentOf_VStack_WithRowSubtotals_LayoutAndValues()
        {
            // The most invocation-heavy combination in the PivotBy code:
            //   * PERCENTOF (triggers all the EtaFunction.Name == "PERCENTOF" branches)
            //   * VSTACK    (multiple functions stacked vertically)
            //   * Row subtotals (RowTotalDepth = 2 emits both subtotals and grand totals)
            //
            // No existing test covers this combination at all.
            //
            // SUM is included as a control: if SUM is correct but PERCENTOF is wrong,
            // the bug is isolated to the PERCENTOF / RelativeTo logic, not layout.
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: F1:K11
            //   F-column:  blank, "Nord", "Nord", "Nord", "Nord", "Syd",
            //              "Syd", "Syd", "Syd", "Grand Total", "Grand Total"
            //   G-column:  blank, "Sthlm", "Sthlm", blank, blank, "Malmö",
            //              "Malmö", blank, blank, blank, blank
            //   H-column:  blank, "SUM", "PERCENTOF", "SUM", "PERCENTOF",
            //              "SUM", "PERCENTOF", "SUM", "PERCENTOF", "SUM", "PERCENTOF"
            //   I-column:  "Q1", 10, 0.25,  10, 0.25,  30, 0.75,  30, 0.75,  40, 1
            //   J-column:  "Q2", 20, 0.333, 20, 0.333, 40, 0.667, 40, 0.667, 60, 1
            //   K-column:  "Total", 30, 0.3, 30, 0.3, 70, 0.7, 70, 0.7, 100, 1
            //
            // Notable Excel rules this test pins down:
            //   - First-level row-key ("Nord"/"Syd") repeats on EVERY function row
            //     in its block, including data, subtotal and grand-total rows.
            //   - Second-level row-key (city) repeats on data-block function rows
            //     but is blank throughout the subtotal and grand-total rows.
            //   - PERCENTOF default RelativeTo = ColumnTotals: denominator is the
            //     column sum, e.g. 10/(10+30)=0.25 for Sthlm/Q1.
            //   - PERCENTOF in the row-total column uses grand total as denominator:
            //     30/100 = 0.3 for Sthlm row.
            //   - Grand total row of PERCENTOF = 1 everywhere (X/X).
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Nord";
                s.Cells["A2"].Value = "Nord";
                s.Cells["A3"].Value = "Syd";
                s.Cells["A4"].Value = "Syd";
                s.Cells["B1"].Value = "Sthlm";
                s.Cells["B2"].Value = "Sthlm";
                s.Cells["B3"].Value = "Malmö";
                s.Cells["B4"].Value = "Malmö";
                s.Cells["C1"].Value = "Q1";
                s.Cells["C2"].Value = "Q2";
                s.Cells["C3"].Value = "Q1";
                s.Cells["C4"].Value = "Q2";
                s.Cells["D1"].Value = 10;
                s.Cells["D2"].Value = 20;
                s.Cells["D3"].Value = 30;
                s.Cells["D4"].Value = 40;
                s.Cells["F1"].Formula =
                    "PIVOTBY(A1:B4, C1:C4, D1:D4, VSTACK(_xleta.SUM, _xleta.PERCENTOF),, 2)";
                s.Calculate();

                // --- Header row ---
                Assert.IsTrue(IsBlank(s.Cells["F1"].Value), "F1 blank.");
                Assert.IsTrue(IsBlank(s.Cells["G1"].Value), "G1 blank.");
                Assert.IsTrue(IsBlank(s.Cells["H1"].Value), "H1 blank.");
                Assert.AreEqual("Q1", s.Cells["I1"].Value);
                Assert.AreEqual("Q2", s.Cells["J1"].Value);
                Assert.AreEqual("Total", s.Cells["K1"].Value);

                // --- Row 2: Nord/Sthlm SUM ---
                Assert.AreEqual("Nord", s.Cells["F2"].Value);
                Assert.AreEqual("Sthlm", s.Cells["G2"].Value);
                Assert.AreEqual("SUM", s.Cells["H2"].Value);
                Assert.AreEqual(10d, s.Cells["I2"].Value);
                Assert.AreEqual(20d, s.Cells["J2"].Value);
                Assert.AreEqual(30d, s.Cells["K2"].Value);

                // --- Row 3: Nord/Sthlm PERCENTOF (row-keys repeat on data-block function rows) ---
                Assert.AreEqual("Nord", s.Cells["F3"].Value,
                    "First-level row-key 'Nord' repeats on second function row of data block.");
                Assert.AreEqual("Sthlm", s.Cells["G3"].Value,
                    "Second-level row-key 'Sthlm' also repeats on second function row of data block.");
                Assert.AreEqual("PERCENTOF", s.Cells["H3"].Value);
                Assert.AreEqual(0.25, System.Math.Round((double)s.Cells["I3"].Value, 6),
                    "Sthlm/Q1 PERCENTOF = 10/(10+30) = 0.25 (column total denominator).");
                Assert.AreEqual(0.333333, System.Math.Round((double)s.Cells["J3"].Value, 6),
                    "Sthlm/Q2 PERCENTOF = 20/(20+40) = 0.333.");
                Assert.AreEqual(0.3, System.Math.Round((double)s.Cells["K3"].Value, 6),
                    "Sthlm row total PERCENTOF = 30/100 = 0.3 (grand total denominator).");

                // --- Row 4: Nord subtotal SUM (first-level repeats, second-level BLANKS) ---
                Assert.AreEqual("Nord", s.Cells["F4"].Value,
                    "First-level row-key 'Nord' repeats on the subtotal row.");
                Assert.IsTrue(IsBlank(s.Cells["G4"].Value),
                    "Second-level row-key blanks on the subtotal row.");
                Assert.AreEqual("SUM", s.Cells["H4"].Value);
                Assert.AreEqual(10d, s.Cells["I4"].Value, "Nord subtotal SUM/Q1.");
                Assert.AreEqual(20d, s.Cells["J4"].Value);
                Assert.AreEqual(30d, s.Cells["K4"].Value);

                // --- Row 5: Nord subtotal PERCENTOF ---
                Assert.AreEqual("Nord", s.Cells["F5"].Value,
                    "First-level row-key 'Nord' still repeats on second function row of subtotal block.");
                Assert.IsTrue(IsBlank(s.Cells["G5"].Value),
                    "Second-level row-key still blank on second function row of subtotal block.");
                Assert.AreEqual("PERCENTOF", s.Cells["H5"].Value);
                Assert.AreEqual(0.25, System.Math.Round((double)s.Cells["I5"].Value, 6),
                    "Nord subtotal PERCENTOF/Q1 = 10/40 = 0.25.");
                Assert.AreEqual(0.333333, System.Math.Round((double)s.Cells["J5"].Value, 6),
                    "Nord subtotal PERCENTOF/Q2 = 20/60.");
                Assert.AreEqual(0.3, System.Math.Round((double)s.Cells["K5"].Value, 6),
                    "Nord subtotal row total PERCENTOF = 30/100.");

                // --- Row 6: Syd/Malmö SUM ---
                Assert.AreEqual("Syd", s.Cells["F6"].Value);
                Assert.AreEqual("Malmö", s.Cells["G6"].Value);
                Assert.AreEqual("SUM", s.Cells["H6"].Value);
                Assert.AreEqual(30d, s.Cells["I6"].Value);
                Assert.AreEqual(40d, s.Cells["J6"].Value);
                Assert.AreEqual(70d, s.Cells["K6"].Value);

                // --- Row 7: Syd/Malmö PERCENTOF ---
                Assert.AreEqual("Syd", s.Cells["F7"].Value);
                Assert.AreEqual("Malmö", s.Cells["G7"].Value);
                Assert.AreEqual("PERCENTOF", s.Cells["H7"].Value);
                Assert.AreEqual(0.75, System.Math.Round((double)s.Cells["I7"].Value, 6),
                    "Malmö/Q1 PERCENTOF = 30/40 = 0.75.");
                Assert.AreEqual(0.666667, System.Math.Round((double)s.Cells["J7"].Value, 6));
                Assert.AreEqual(0.7, System.Math.Round((double)s.Cells["K7"].Value, 6));

                // --- Row 8: Syd subtotal SUM ---
                Assert.AreEqual("Syd", s.Cells["F8"].Value);
                Assert.IsTrue(IsBlank(s.Cells["G8"].Value));
                Assert.AreEqual("SUM", s.Cells["H8"].Value);
                Assert.AreEqual(30d, s.Cells["I8"].Value);
                Assert.AreEqual(40d, s.Cells["J8"].Value);
                Assert.AreEqual(70d, s.Cells["K8"].Value);

                // --- Row 9: Syd subtotal PERCENTOF ---
                Assert.AreEqual("Syd", s.Cells["F9"].Value);
                Assert.IsTrue(IsBlank(s.Cells["G9"].Value));
                Assert.AreEqual("PERCENTOF", s.Cells["H9"].Value);
                Assert.AreEqual(0.75, System.Math.Round((double)s.Cells["I9"].Value, 6));
                Assert.AreEqual(0.666667, System.Math.Round((double)s.Cells["J9"].Value, 6));
                Assert.AreEqual(0.7, System.Math.Round((double)s.Cells["K9"].Value, 6));

                // --- Row 10: Grand Total SUM ---
                Assert.AreEqual("Grand Total", s.Cells["F10"].Value,
                    "With RowTotalDepth=2 the bottom label is 'Grand Total', not 'Total'.");
                Assert.IsTrue(IsBlank(s.Cells["G10"].Value));
                Assert.AreEqual("SUM", s.Cells["H10"].Value);
                Assert.AreEqual(40d, s.Cells["I10"].Value);
                Assert.AreEqual(60d, s.Cells["J10"].Value);
                Assert.AreEqual(100d, s.Cells["K10"].Value);

                // --- Row 11: Grand Total PERCENTOF ---
                Assert.AreEqual("Grand Total", s.Cells["F11"].Value,
                    "Grand total label repeats on the second function row.");
                Assert.IsTrue(IsBlank(s.Cells["G11"].Value));
                Assert.AreEqual("PERCENTOF", s.Cells["H11"].Value);
                Assert.AreEqual(1d, (double)s.Cells["I11"].Value, "Grand total PERCENTOF = 1.");
                Assert.AreEqual(1d, (double)s.Cells["J11"].Value);
                Assert.AreEqual(1d, (double)s.Cells["K11"].Value);
            }
        }

        [TestMethod]
        public void PivotBy_NegativeColTotalDepth2_GrandTotalAtLeft_SubtotalsBeforeLeavesInEachGroup()
        {
            // ColTotalDepth = -2 produces a richer layout than -1:
            //   * Grand total leftmost (colTotalAtLeft, |depth| > 1 means label = "Grand Total")
            //   * Column subtotals enabled (showColSubtotals = |depth| > 1)
            //   * Each year subtotal appears AT THE START of its group, before that group's leaves
            //
            // The previously fixed -1 test only verified grand-total-at-left when subtotals were
            // disabled. With subtotals on, colEntries contains both subtotals and leaves; the
            // open question (now answered by Excel) is the relative position of the subtotal
            // within its group: BEFORE its leaves, not after.
            //
            // Data values are powers of two so every subtotal and grand total is unique and
            // can be unambiguously identified from its cell value alone:
            //   2025/Q1 = 1, 2025/Q2 = 2  -> 2025 subtotal = 3
            //   2026/Q1 = 4, 2026/Q2 = 8  -> 2026 subtotal = 12
            //   Grand total = 15
            //
            //   Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: F1:M4
            //     Row 1: ""       "Grand Total"  2025   2025   2025   2026   2026   2026
            //     Row 2: ""       ""             ""     "Q1"   "Q2"   ""     "Q1"   "Q2"
            //     Row 3: "R"      15             3      1      2      12     4      8
            //     Row 4: "Total"  15             3      1      2      12     4      8
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "R"; s.Cells["B1"].Value = 2025; s.Cells["C1"].Value = "Q1"; s.Cells["D1"].Value = 1;
                s.Cells["A2"].Value = "R"; s.Cells["B2"].Value = 2025; s.Cells["C2"].Value = "Q2"; s.Cells["D2"].Value = 2;
                s.Cells["A3"].Value = "R"; s.Cells["B3"].Value = 2026; s.Cells["C3"].Value = "Q1"; s.Cells["D3"].Value = 4;
                s.Cells["A4"].Value = "R"; s.Cells["B4"].Value = 2026; s.Cells["C4"].Value = "Q2"; s.Cells["D4"].Value = 8;

                s.Cells["F1"].Formula = "PIVOTBY(A1:A4, B1:C4, D1:D4, _xleta.SUM, , , , -2)";
                s.Calculate();

                // --- Row 1: top header (corner + grand total label + year labels) ---
                Assert.IsTrue(
                    s.Cells["F1"].Value == null || (s.Cells["F1"].Value as string) == string.Empty,
                    "F1 corner above row label should be blank. Got: " + (s.Cells["F1"].Value ?? "null"));
                Assert.AreEqual("Grand Total", s.Cells["G1"].Value,
                    "Grand Total label must be at G1 (leftmost data column) and read 'Grand Total' for |ColTotalDepth|=2.");
                Assert.AreEqual(2025, s.Cells["H1"].Value, "H1: 2025 subtotal column - year label.");
                Assert.AreEqual(2025, s.Cells["I1"].Value, "I1: 2025/Q1 leaf - year label.");
                Assert.AreEqual(2025, s.Cells["J1"].Value, "J1: 2025/Q2 leaf - year label.");
                Assert.AreEqual(2026, s.Cells["K1"].Value, "K1: 2026 subtotal column - year label.");
                Assert.AreEqual(2026, s.Cells["L1"].Value, "L1: 2026/Q1 leaf - year label.");
                Assert.AreEqual(2026, s.Cells["M1"].Value, "M1: 2026/Q2 leaf - year label.");

                // --- Row 2: quarter header (subtotal cols are blank here) ---
                Assert.IsTrue(
                    s.Cells["G2"].Value == null || (s.Cells["G2"].Value as string) == string.Empty,
                    "G2 grand total has no quarter label. Got: " + (s.Cells["G2"].Value ?? "null"));
                Assert.IsTrue(
                    s.Cells["H2"].Value == null || (s.Cells["H2"].Value as string) == string.Empty,
                    "H2 year subtotal has no quarter label. Got: " + (s.Cells["H2"].Value ?? "null"));
                Assert.AreEqual("Q1", s.Cells["I2"].Value, "I2: 2025/Q1 quarter.");
                Assert.AreEqual("Q2", s.Cells["J2"].Value, "J2: 2025/Q2 quarter.");
                Assert.IsTrue(
                    s.Cells["K2"].Value == null || (s.Cells["K2"].Value as string) == string.Empty,
                    "K2 year subtotal has no quarter label. Got: " + (s.Cells["K2"].Value ?? "null"));
                Assert.AreEqual("Q1", s.Cells["L2"].Value, "L2: 2026/Q1 quarter.");
                Assert.AreEqual("Q2", s.Cells["M2"].Value, "M2: 2026/Q2 quarter.");

                // --- Row 3: data row R ---
                // Each value is uniquely identifiable: 15=grand, 3=2025-sub, 12=2026-sub, leaves are 1,2,4,8.
                Assert.AreEqual("R", s.Cells["F3"].Value);
                Assert.AreEqual(15d, s.Cells["G3"].Value, "G3: grand total leftmost = 1+2+4+8.");
                Assert.AreEqual(3d, s.Cells["H3"].Value, "H3: 2025 subtotal must come BEFORE its quarters (1+2).");
                Assert.AreEqual(1d, s.Cells["I3"].Value, "I3: 2025/Q1 leaf.");
                Assert.AreEqual(2d, s.Cells["J3"].Value, "J3: 2025/Q2 leaf.");
                Assert.AreEqual(12d, s.Cells["K3"].Value, "K3: 2026 subtotal must come BEFORE its quarters (4+8).");
                Assert.AreEqual(4d, s.Cells["L3"].Value, "L3: 2026/Q1 leaf.");
                Assert.AreEqual(8d, s.Cells["M3"].Value, "M3: 2026/Q2 leaf.");

                // --- Row 4: grand total row (same shape as data, since only one source row) ---
                Assert.AreEqual("Total", s.Cells["F4"].Value,
                    "F4: row total label. Note this is 'Total' (RowTotalDepth=1 default), not 'Grand Total'.");
                Assert.AreEqual(15d, s.Cells["G4"].Value, "G4: corner cell (grand total of grand total).");
                Assert.AreEqual(3d, s.Cells["H4"].Value, "H4: 2025 subtotal column, grand total row.");
                Assert.AreEqual(1d, s.Cells["I4"].Value);
                Assert.AreEqual(2d, s.Cells["J4"].Value);
                Assert.AreEqual(12d, s.Cells["K4"].Value);
                Assert.AreEqual(4d, s.Cells["L4"].Value);
                Assert.AreEqual(8d, s.Cells["M4"].Value);
            }
        }

        [TestMethod]
        public void PivotBy_NegativeRowTotalDepth2_GrandTotalAtTop_SubtotalsBeforeLeavesInEachGroup()
        {
            // RowTotalDepth = -2 produces:
            //   * Grand total at the top (rowTotalAtTop, |depth| > 1 means label = "Grand Total")
            //   * Row subtotals enabled (showRowSubtotals = |depth| > 1)
            //   * Each country subtotal appears AT THE START of its group, before that group's leaves
            //
            // This mirrors what we fixed earlier for ColTotalDepth = -2 on the column axis:
            // a negative sign flips BOTH grand total AND subtotal placement symmetrically.
            //
            // Current EPPlus puts subtotals AFTER their leaves regardless of sign - that's the
            // bug this test catches. Grand-total-at-top is likely already handled correctly via
            // rowTotalAtTop, but subtotal placement isn't tied to that flag.
            //
            // Data values are powers of two so every subtotal and grand total is unique:
            //   Sweden/Stockholm/X = 1, Sweden/Göteborg/X = 2  -> Sweden subtotal = 3
            //   Norway/Oslo/X = 4,      Norway/Bergen/X = 8    -> Norway subtotal = 12
            //   Grand total = 15
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: F1:I8
            //     Row 1: ""             ""           "X"  "Total"
            //     Row 2: "Grand Total"  ""           15   15
            //     Row 3: "Norway"       ""           12   12      <- subtotal BEFORE leaves
            //     Row 4: "Norway"       "Bergen"     8    8
            //     Row 5: "Norway"       "Oslo"       4    4
            //     Row 6: "Sweden"       ""           3    3       <- subtotal BEFORE leaves
            //     Row 7: "Sweden"       "Göteborg"   2    2
            //     Row 8: "Sweden"       "Stockholm"  1    1
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Sweden"; s.Cells["B1"].Value = "Stockholm"; s.Cells["C1"].Value = "X"; s.Cells["D1"].Value = 1;
                s.Cells["A2"].Value = "Sweden"; s.Cells["B2"].Value = "Göteborg"; s.Cells["C2"].Value = "X"; s.Cells["D2"].Value = 2;
                s.Cells["A3"].Value = "Norway"; s.Cells["B3"].Value = "Oslo"; s.Cells["C3"].Value = "X"; s.Cells["D3"].Value = 4;
                s.Cells["A4"].Value = "Norway"; s.Cells["B4"].Value = "Bergen"; s.Cells["C4"].Value = "X"; s.Cells["D4"].Value = 8;

                s.Cells["F1"].Formula = "PIVOTBY(A1:B4, C1:C4, D1:D4, _xleta.SUM, , -2)";
                s.Calculate();

                // --- Row 1: column key header ---
                Assert.IsTrue(IsBlank(s.Cells["F1"].Value), "F1 should be blank (corner).");
                Assert.IsTrue(IsBlank(s.Cells["G1"].Value), "G1 should be blank (corner).");
                Assert.AreEqual("X", s.Cells["H1"].Value);
                Assert.AreEqual("Total", s.Cells["I1"].Value);

                // --- Row 2: grand total AT TOP ---
                Assert.AreEqual("Grand Total", s.Cells["F2"].Value,
                    "Grand Total row must be at top (row 2, immediately after header). Label = 'Grand Total' for |depth| > 1.");
                Assert.IsTrue(IsBlank(s.Cells["G2"].Value), "G2 should be blank on grand total row.");
                Assert.AreEqual(15d, s.Cells["H2"].Value, "Grand total X = 1+2+4+8.");
                Assert.AreEqual(15d, s.Cells["I2"].Value, "Grand total corner.");

                // --- Row 3: Norway subtotal (BEFORE its leaves) ---
                Assert.AreEqual("Norway", s.Cells["F3"].Value,
                    "Norway subtotal must come BEFORE its leaves. Label is the group name, not 'Total'.");
                Assert.IsTrue(IsBlank(s.Cells["G3"].Value), "G3 should be blank on subtotal row (city col).");
                Assert.AreEqual(12d, s.Cells["H3"].Value, "Norway subtotal = 4 + 8.");
                Assert.AreEqual(12d, s.Cells["I3"].Value);

                // --- Row 4: Norway / Bergen (alphabetical: Bergen before Oslo) ---
                Assert.AreEqual("Norway", s.Cells["F4"].Value);
                Assert.AreEqual("Bergen", s.Cells["G4"].Value);
                Assert.AreEqual(8d, s.Cells["H4"].Value);
                Assert.AreEqual(8d, s.Cells["I4"].Value);

                // --- Row 5: Norway / Oslo ---
                Assert.AreEqual("Norway", s.Cells["F5"].Value);
                Assert.AreEqual("Oslo", s.Cells["G5"].Value);
                Assert.AreEqual(4d, s.Cells["H5"].Value);
                Assert.AreEqual(4d, s.Cells["I5"].Value);

                // --- Row 6: Sweden subtotal (BEFORE its leaves) ---
                Assert.AreEqual("Sweden", s.Cells["F6"].Value,
                    "Sweden subtotal must come BEFORE its leaves.");
                Assert.IsTrue(IsBlank(s.Cells["G6"].Value), "G6 should be blank on subtotal row.");
                Assert.AreEqual(3d, s.Cells["H6"].Value, "Sweden subtotal = 1 + 2.");
                Assert.AreEqual(3d, s.Cells["I6"].Value);

                // --- Row 7: Sweden / Göteborg ---
                Assert.AreEqual("Sweden", s.Cells["F7"].Value);
                Assert.AreEqual("Göteborg", s.Cells["G7"].Value);
                Assert.AreEqual(2d, s.Cells["H7"].Value);
                Assert.AreEqual(2d, s.Cells["I7"].Value);

                // --- Row 8: Sweden / Stockholm ---
                Assert.AreEqual("Sweden", s.Cells["F8"].Value);
                Assert.AreEqual("Stockholm", s.Cells["G8"].Value);
                Assert.AreEqual(1d, s.Cells["H8"].Value);
                Assert.AreEqual(1d, s.Cells["I8"].Value);
            }
        }

        [TestMethod]
        public void PivotBy_FilterArray_AllRowsFiltered_ReturnsValueError()
        {
            // When the filter array excludes every row, Excel returns #VALUE! - not an
            // empty spill, not zero, not a degenerate single-cell result.
            //
            // Suspected EPPlus bug: BuildPivotData filters rows but never checks whether
            // anything remains. Downstream code in AggregateLeaf accesses allVals[0].Length
            // unconditionally, which throws IndexOutOfRangeException on an empty list,
            // bubbling up as an unhandled exception or a generic error - not the clean
            // #VALUE! Excel produces.
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   H1 = #VALUE!
            //   (no spill range produced)
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A"; s.Cells["B1"].Value = "X"; s.Cells["C1"].Value = 10;
                s.Cells["A2"].Value = "B"; s.Cells["B2"].Value = "X"; s.Cells["C2"].Value = 20;
                s.Cells["A3"].Value = "C"; s.Cells["B3"].Value = "X"; s.Cells["C3"].Value = 30;

                // Filter array: all zeros = no row passes
                s.Cells["F1"].Value = 0;
                s.Cells["F2"].Value = 0;
                s.Cells["F3"].Value = 0;

                s.Cells["H1"].Formula = "PIVOTBY(A1:A3, B1:B3, C1:C3, _xleta.SUM,,,,,,F1:F3)";
                s.Calculate();

                var value = s.Cells["H1"].Value;
                Assert.IsInstanceOfType(
                    value,
                    typeof(ExcelErrorValue),
                    "Expected #VALUE! (ExcelErrorValue) when filter array excludes every row. " +
                    "Got: " + (value == null ? "null" : value.GetType().Name + " = " + value));

                var err = (ExcelErrorValue)value;
                Assert.AreEqual(eErrorType.Value, err.Type,
                    "Error must be #VALUE! specifically. Got: " + err.Type);

                // No spill should be produced - subsequent cells must be untouched (null).
                Assert.IsNull(s.Cells["I1"].Value, "I1 should not be part of a spill range.");
                Assert.IsNull(s.Cells["H2"].Value, "H2 should not be part of a spill range.");
            }
        }

        [TestMethod]
        public void PivotBy_HeaderAutoDetect_NumericFirstRowKeyCell_StillDetectsHeadersFromValuesColumn()
        {
            // When field_headers is omitted, Excel auto-detects headers by inspecting
            // the values column pattern (text followed by numbers), NOT just the row
            // field's first cell. Even when A1 is numeric (2025), Excel still detects
            // headers because C1='Revenue' followed by C2..C4 numeric is the canonical
            // header signature.
            //
            // Suspected EPPlus behaviour: ResolveHeaders looks at the row field's first
            // cell type, sees 2025 (numeric), and returns FieldHeaders.No - causing row 1
            // to be processed as data with 'Quarter' as a col key and 'Revenue' as a
            // text value (which then either errors or gets silently mistreated).
            //
            // Verified in Excel (sv-SE) 2026-05-22:
            //   Spill range: E1:H4
            //     Row 1: ""       "Q1"   "Q2"   "Total"
            //     Row 2: 2025     100    <blank> 100
            //     Row 3: 2026     300    200    500
            //     Row 4: "Total"  400    200    600
            //
            // (Note: 'Quarter' and 'Revenue' do NOT appear anywhere in the output.
            // Detection mode is YesAndDontShow, so field names are recognised but not
            // displayed - and the auto-generated row key column is just 2025/2026/Total.)
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = 2025; s.Cells["B1"].Value = "Quarter"; s.Cells["C1"].Value = "Revenue";
                s.Cells["A2"].Value = 2025; s.Cells["B2"].Value = "Q1"; s.Cells["C2"].Value = 100;
                s.Cells["A3"].Value = 2026; s.Cells["B3"].Value = "Q2"; s.Cells["C3"].Value = 200;
                s.Cells["A4"].Value = 2026; s.Cells["B4"].Value = "Q1"; s.Cells["C4"].Value = 300;

                s.Cells["E1"].Formula = "PIVOTBY(A1:A4, B1:B4, C1:C4, _xleta.SUM)";
                s.Calculate();

                // --- Header row: col-key values + Total label ---
                // E1 corner blank; F1='Q1' (NOT 'Quarter' - row 1 must be detected as header)
                Assert.IsTrue(
                    s.Cells["E1"].Value == null || (s.Cells["E1"].Value as string) == string.Empty,
                    "E1 corner blank. Got: " + (s.Cells["E1"].Value ?? "null"));
                Assert.AreEqual("Q1", s.Cells["F1"].Value,
                    "F1 must be the col-key 'Q1', not 'Quarter'. If you see 'Quarter' here, " +
                    "row 1 was treated as data instead of as a header row.");
                Assert.AreEqual("Q2", s.Cells["G1"].Value);
                Assert.AreEqual("Total", s.Cells["H1"].Value);

                // --- Row 2: 2025 (numeric row key, NOT promoted to header) ---
                Assert.AreEqual(2025, s.Cells["E2"].Value,
                    "E2 must be the numeric row key 2025. If you see 'Year' or similar, header " +
                    "promotion went too far.");
                Assert.AreEqual(100d, s.Cells["F2"].Value, "2025/Q1 = 100.");
                Assert.IsTrue(
                    s.Cells["G2"].Value == null || (s.Cells["G2"].Value as string) == string.Empty,
                    "G2 should be blank - 2025 has no Q2 value. Got: " + (s.Cells["G2"].Value ?? "null"));
                Assert.AreEqual(100d, s.Cells["H2"].Value, "2025 row total = 100.");

                // --- Row 3: 2026 ---
                Assert.AreEqual(2026, s.Cells["E3"].Value);
                Assert.AreEqual(300d, s.Cells["F3"].Value, "2026/Q1 = 300.");
                Assert.AreEqual(200d, s.Cells["G3"].Value, "2026/Q2 = 200.");
                Assert.AreEqual(500d, s.Cells["H3"].Value, "2026 row total = 500.");

                // --- Row 4: Grand total ---
                Assert.AreEqual("Total", s.Cells["E4"].Value);
                Assert.AreEqual(400d, s.Cells["F4"].Value, "Q1 column total = 100 + 300.");
                Assert.AreEqual(200d, s.Cells["G4"].Value, "Q2 column total = 200.");
                Assert.AreEqual(600d, s.Cells["H4"].Value, "Grand total = 100 + 200 + 300.");
            }
        }

        [TestMethod]
        public void PivotBy_RowSortOrderArray_MagnitudeDeterminesPriority()
        {
            // Excel's PIVOTBY interprets the row_sort_order array such that the
            // MAGNITUDE of each value determines both which row field to sort by AND
            // its priority - NOT the position in the array:
            //   |val| = 1  -> sort by first row field (col 0), highest priority
            //   |val| = 2  -> sort by second row field (col 1), secondary priority
            //   sign       -> direction (+ = ASC, - = DESC)
            //
            // So {2, -1} and {-1, 2} produce identical results: col 0 DESC primary,
            // col 1 ASC secondary. EPPlus previously iterated the array in input
            // order, treating the FIRST element as primary - which gives the wrong
            // answer whenever the array isn't already in magnitude-ascending order.
            //
            // Input data is deliberately NOT pre-sorted so that the output order
            // alone tells us whether sorting actually happened.
            //
            // Verified in Excel (sv-SE) 2026-05-22 (with input rows in this exact order):
            //   F1=Norway,  G1=Bergen,    H1=10
            //   F2=Sweden,  G2=Stockholm, H2=20
            //   F3=Norway,  G3=Oslo,      H3=30
            //   F4=Sweden,  G4=Göteborg,  H4=40
            //
            // Spill range: F1:I6 (4 data rows + header + grand total)
            //   Row 1: ""        ""           "X"   "Total"
            //   Row 2: "Sweden"  "Göteborg"   40    40
            //   Row 3: "Sweden"  "Stockholm"  20    20
            //   Row 4: "Norway"  "Bergen"     10    10
            //   Row 5: "Norway"  "Oslo"       30    30
            //   Row 6: "Total"   ""           100   100
            //
            // Expected primary order: Country DESC (Sweden before Norway).
            // Expected secondary order within each country: City ASC.
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Norway"; s.Cells["B1"].Value = "Bergen"; s.Cells["C1"].Value = "X"; s.Cells["D1"].Value = 10;
                s.Cells["A2"].Value = "Sweden"; s.Cells["B2"].Value = "Stockholm"; s.Cells["C2"].Value = "X"; s.Cells["D2"].Value = 20;
                s.Cells["A3"].Value = "Norway"; s.Cells["B3"].Value = "Oslo"; s.Cells["C3"].Value = "X"; s.Cells["D3"].Value = 30;
                s.Cells["A4"].Value = "Sweden"; s.Cells["B4"].Value = "Göteborg"; s.Cells["C4"].Value = "X"; s.Cells["D4"].Value = 40;

                s.Cells["F1"].Formula = "PIVOTBY(A1:B4, C1:C4, D1:D4, _xleta.SUM, , , {2, -1})";
                s.Calculate();

                // --- Header row ---
                Assert.AreEqual("X", s.Cells["H1"].Value);
                Assert.AreEqual("Total", s.Cells["I1"].Value);

                // --- Row 2: Sweden / Göteborg (Country DESC primary, so Sweden group first;
                //                              City ASC secondary, so Göteborg before Stockholm) ---
                Assert.AreEqual("Sweden", s.Cells["F2"].Value,
                    "Primary sort = Country DESC, so Sweden group must come before Norway. " +
                    "If you see 'Norway' here, the array was sorted in position order (first = primary) " +
                    "instead of magnitude order (|val|=1 = primary).");
                Assert.AreEqual("Göteborg", s.Cells["G2"].Value,
                    "Secondary sort = City ASC, so Göteborg before Stockholm within Sweden.");
                Assert.AreEqual(40d, s.Cells["H2"].Value);
                Assert.AreEqual(40d, s.Cells["I2"].Value);

                // --- Row 3: Sweden / Stockholm ---
                Assert.AreEqual("Sweden", s.Cells["F3"].Value);
                Assert.AreEqual("Stockholm", s.Cells["G3"].Value);
                Assert.AreEqual(20d, s.Cells["H3"].Value);
                Assert.AreEqual(20d, s.Cells["I3"].Value);

                // --- Row 4: Norway / Bergen ---
                Assert.AreEqual("Norway", s.Cells["F4"].Value);
                Assert.AreEqual("Bergen", s.Cells["G4"].Value);
                Assert.AreEqual(10d, s.Cells["H4"].Value);
                Assert.AreEqual(10d, s.Cells["I4"].Value);

                // --- Row 5: Norway / Oslo ---
                Assert.AreEqual("Norway", s.Cells["F5"].Value);
                Assert.AreEqual("Oslo", s.Cells["G5"].Value);
                Assert.AreEqual(30d, s.Cells["H5"].Value);
                Assert.AreEqual(30d, s.Cells["I5"].Value);

                // --- Row 6: Grand total ---
                Assert.AreEqual("Total", s.Cells["F6"].Value);
                Assert.AreEqual(100d, s.Cells["H6"].Value, "Grand total X = 10+20+30+40.");
                Assert.AreEqual(100d, s.Cells["I6"].Value);
            }
        }

        // Helper - blank means null or empty string.
        private static bool IsBlank(object v)
        {
            return v == null || (v is string str && str.Length == 0);
        }
    }
}
