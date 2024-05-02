using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class MInverseTests : TestBase
    {
        [TestMethod]
        public void MInverseTest()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 7;
            ws.Cells["B1"].Value = 2;
            ws.Cells["C1"].Value = 1;
            ws.Cells["A2"].Value = 0;
            ws.Cells["B2"].Value = 3;
            ws.Cells["C2"].Value = -1;
            ws.Cells["A3"].Value = -3;
            ws.Cells["B3"].Value = 4;
            ws.Cells["C3"].Value = 2;
            ws.Cells["E1"].Formula = "MINVERSE(A1:C3)";
            ws.Calculate();
            Assert.AreEqual(0.117647059d, System.Math.Round((double)ws.Cells["E1"].Value, 9));
            Assert.AreEqual(0d, System.Math.Round((double)ws.Cells["F1"].Value, 9));
            Assert.AreEqual(-0.058823529d, System.Math.Round((double)ws.Cells["G1"].Value, 9));
            Assert.AreEqual(0.035294118d, System.Math.Round((double)ws.Cells["E2"].Value, 9));
            Assert.AreEqual(0.2d, System.Math.Round((double)ws.Cells["F2"].Value, 9));
            Assert.AreEqual(0.082352941d, System.Math.Round((double)ws.Cells["G2"].Value, 9));
            Assert.AreEqual(0.105882353d, System.Math.Round((double)ws.Cells["E3"].Value, 9));
            Assert.AreEqual(-0.4d, System.Math.Round((double)ws.Cells["F3"].Value, 9));
            Assert.AreEqual(0.247058824d, System.Math.Round((double)ws.Cells["G3"].Value, 9));
        }

        [TestMethod]
        public void MInverseNonSquareTest()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 4;
            ws.Cells["C1"].Value = 7;
            ws.Cells["A2"].Value = 2;
            ws.Cells["B2"].Value = 5;
            ws.Cells["C2"].Value = 8;
            ws.Cells["E1"].Formula = "MINVERSE(A1:C3)";
            ws.Cells["H1"].Formula = "MINVERSE(A1:C2)";
            ws.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E1"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["H1"].Value);
        }

        [TestMethod]
        public void MInverseArrayTest()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["E1"].Formula = "MINVERSE({7,2,1;0,3,-1;-3,4,2})";
            ws.Calculate();
            Assert.AreEqual(0.117647059d, System.Math.Round((double)ws.Cells["E1"].Value, 9));
            Assert.AreEqual(0d, System.Math.Round((double)ws.Cells["F1"].Value, 9));
            Assert.AreEqual(-0.058823529d, System.Math.Round((double)ws.Cells["G1"].Value, 9));
            Assert.AreEqual(0.035294118d, System.Math.Round((double)ws.Cells["E2"].Value, 9));
            Assert.AreEqual(0.2d, System.Math.Round((double)ws.Cells["F2"].Value, 9));
            Assert.AreEqual(0.082352941d, System.Math.Round((double)ws.Cells["G2"].Value, 9));
            Assert.AreEqual(0.105882353d, System.Math.Round((double)ws.Cells["E3"].Value, 9));
            Assert.AreEqual(-0.4d, System.Math.Round((double)ws.Cells["F3"].Value, 9));
            Assert.AreEqual(0.247058824d, System.Math.Round((double)ws.Cells["G3"].Value, 9));
        }

        [TestMethod]
        public void MInverseDET0Test()
        {
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 1;
            ws.Cells["B1"].Value = 4;
            ws.Cells["C1"].Value = 7;
            ws.Cells["A2"].Value = 2;
            ws.Cells["B2"].Value = 5;
            ws.Cells["C2"].Value = 8;
            ws.Cells["A3"].Value = 3;
            ws.Cells["B3"].Value = 6;
            ws.Cells["C3"].Value = 9;
            ws.Cells["E1"].Formula = "MINVERSE(A1:C3)";
            ws.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), ws.Cells["E1"].Value);
        }

        [TestMethod]
        public void MInverseWorkbookTest()
        {
            using var p = OpenPackage("MInverseTest.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 7;
            ws.Cells["B1"].Value = 2;
            ws.Cells["C1"].Value = 1;
            ws.Cells["A2"].Value = 0;
            ws.Cells["B2"].Value = 3;
            ws.Cells["C2"].Value = -1;
            ws.Cells["A3"].Value = -3;
            ws.Cells["B3"].Value = 4;
            ws.Cells["C3"].Value = 2;
            ws.Cells["E1"].Formula = "MINVERSE(A1:C3)";
            ws.Cells["A5"].Value = 1;
            ws.Cells["B5"].Value = 4;
            ws.Cells["C5"].Value = 7;
            ws.Cells["A6"].Value = 2;
            ws.Cells["B6"].Value = 5;
            ws.Cells["C6"].Value = 8;
            ws.Cells["E6"].Formula = "MINVERSE(A5:C7)";
            ws.Cells["A10"].Value = 1;
            ws.Cells["B10"].Value = 4;
            ws.Cells["C10"].Value = 7;
            ws.Cells["A11"].Value = 2;
            ws.Cells["B11"].Value = 5;
            ws.Cells["C11"].Value = 8;
            ws.Cells["A12"].Value = 3;
            ws.Cells["B12"].Value = 6;
            ws.Cells["C12"].Value = 9;
            ws.Cells["E10"].Formula = "MINVERSE(A10:C12)";
            ws.Calculate();
            SaveAndCleanup(p);
        }

        //Determinant Test
        [TestMethod]
        public void DeterminantTest()
        {
            using var p = OpenPackage("MDetermTest.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 7;
            ws.Cells["B1"].Value = 2;
            ws.Cells["C1"].Value = 1;
            ws.Cells["A2"].Value = 0;
            ws.Cells["B2"].Value = 3;
            ws.Cells["C2"].Value = -1;
            ws.Cells["A3"].Value = -3;
            ws.Cells["B3"].Value = 4;
            ws.Cells["C3"].Value = 2;
            ws.Cells["E1"].Formula = "MDETERM(A1:C3)";
            ws.Cells["A5"].Value = 1;
            ws.Cells["B5"].Value = 4;
            ws.Cells["C5"].Value = 7;
            ws.Cells["A6"].Value = 2;
            ws.Cells["B6"].Value = 5;
            ws.Cells["C6"].Value = 8;
            ws.Cells["E6"].Formula = "MDETERM(A5:C7)";
            ws.Cells["A10"].Value = 1;
            ws.Cells["B10"].Value = 4;
            ws.Cells["C10"].Value = 7;
            ws.Cells["A11"].Value = 2;
            ws.Cells["B11"].Value = 5;
            ws.Cells["C11"].Value = 8;
            ws.Cells["A12"].Value = 3;
            ws.Cells["B12"].Value = 6;
            ws.Cells["C12"].Value = 9;
            ws.Cells["E10"].Formula = "MDETERM(A10:C12)";
            ws.Cells["E15"].Formula = "MDETERM({7,2,1;0,3,-1;-3,4,2})";

            ws.Calculate();
            Assert.AreEqual(85d, ws.Cells["E1"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E6"].Value);
            Assert.AreEqual(0d, ws.Cells["E10"].Value);
            Assert.AreEqual(85d, ws.Cells["E15"].Value);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void DeterminantLinestDataTest()
        {
            using var p = OpenTemplatePackage(@"Linest_EPPlusTest.xlsx");
            var ws = p.Workbook.Worksheets["EPPLUS Resultat"];
            ws.Cells["Z1"].Formula = "MDETERM(B4:F9)";
            ws.Cells["Z2"].Formula = "MDETERM(I4:L7)";
            ws.Cells["Z3"].Formula = "MDETERM(O4:R8)";
            ws.Cells["Z4"].Formula = "MDETERM(B24:F25)";
            ws.Cells["Z5"].Formula = "MDETERM(I24:M33)";
            ws.Cells["Z6"].Formula = "MDETERM(P24:T33)";
            ws.Cells["Z7"].Formula = "MDETERM(B50:D54)";
            ws.Cells["Z8"].Formula = "MDETERM(G50:I54)";
            ws.Cells["Z9"].Formula = "MDETERM(K50:N54)";
            ws.Cells["Z10"].Formula = "MDETERM(P50:S53)";
            ws.Cells["Z11"].Formula = "MDETERM(B72:L1071)";
            ws.Cells["Z12"].Formula = "MDETERM(N88:T95)";
            ws.Calculate();
            SaveAndCleanup (p);
        }

        //MUnit Test

        [TestMethod]
        public void MatrixUnitTest()
        {
            using var p = OpenPackage("MUnitTest.xlsx", true);
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.Cells["A1"].Value = 7;
            ws.Cells["A2"].Value = 2;
            ws.Cells["A3"].Value = 1;
            ws.Cells["A4"].Value = 0;
            ws.Cells["A5"].Value = -4;
            ws.Cells["A6"].Value = "k";
            ws.Cells["A7"].Value = 2.2d;

            ws.Cells["E1"].Formula = "MUNIT(A1)";
            ws.Cells["E10"].Formula = "MUNIT(A2)";
            ws.Cells["E15"].Formula = "MUNIT(A3)";
            ws.Cells["E18"].Formula = "MUNIT(A4)";
            ws.Cells["E20"].Formula = "MUNIT(A5)";
            ws.Cells["E22"].Formula = "MUNIT(A6)";
            ws.Cells["E30"].Formula = "MUNIT(A6:A7)";
            ws.Cells["A10"].Formula = "MUNIT(3)";
            ws.Cells["E25"].Formula = "MUNIT(A7)";
            ws.Calculate();
            Assert.AreEqual(1d, ws.Cells["E1"].Value);
            Assert.AreEqual(1d, ws.Cells["E10"].Value);
            Assert.AreEqual(1d, ws.Cells["E15"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E18"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E20"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E22"].Value);
            Assert.AreEqual(1d, ws.Cells["A10"].Value);
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ws.Cells["E30"].Value);
            Assert.AreEqual(1d, ws.Cells["E31"].Value);
            SaveAndCleanup(p);
        }
    }
}
