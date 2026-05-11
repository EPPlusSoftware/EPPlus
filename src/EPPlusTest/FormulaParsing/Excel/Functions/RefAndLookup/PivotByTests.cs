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
                //Fixa detta test det är förskjutet en rad fel
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
                s.Cells["E1"].Formula = "PIVOTBY(A1:A3,B1:C3,D1:D3,_xleta.PERCENTOF,,,,,,,3)";
                s.Calculate();
            }
        }

        [TestMethod]
        public void PivotByRelativeTo2()
        {
            using (var package = new ExcelPackage())
            {
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
                var sheet = package.Workbook.Worksheets[2];
                package.Workbook.CalcMode = ExcelCalcMode.Manual;

                sheet.Cells["B17"].Formula = "PIVOTBY('FCL V'!C6:D2055,'FCL V'!Y6:Y2055,'FCL V'!DH6:DH2055, _xleta.SUM)";
                //sheet.Calculate();
                sheet.Cells["B17"].Calculate();

                Assert.AreEqual("Albania", sheet.Cells["D17"].Value);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void PivotByTemplateTest3()
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
        public void PivotByTemplateTest4()
        {
            using (var package = OpenTemplatePackage("PivotByTest1.xlsx"))
            {
                var sheet = package.Workbook.Worksheets[3];
                package.Workbook.CalcMode = ExcelCalcMode.Manual;

                sheet.Cells["B26"].Formula = "PIVOTBY('FCL V'!C6:D2055,'FCL V'!Y6:Y2055,'FCL V'!DH6:DH2055, _xleta.PERCENTOF,,2,,,,,)";
                sheet.Cells["B26"].Calculate();

                Assert.AreEqual("Albania", sheet.Cells["D26"].Value);
                Assert.AreEqual(0.99237652, System.Math.Round((double)sheet.Cells["G27"].Value), 8d); 
                //SaveAndCleanup(package);
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
    }
}
