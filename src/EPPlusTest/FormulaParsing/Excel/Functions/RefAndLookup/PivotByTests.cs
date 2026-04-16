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
            }
        }

    }
}
