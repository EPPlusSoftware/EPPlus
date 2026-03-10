using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;


namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GroupByTests
    {

        [TestMethod]
        public void GroupBy()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["A3"].Value = "Bertil";
                s.Cells["A4"].Value = "Joe";
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["B3"].Value = 3;
                s.Cells["B4"].Value = 0;
                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, _xleta.SUM)";
                s.Calculate();
                Assert.AreEqual("Anna", s.Cells["C1"].Value);
                Assert.AreEqual("Bertil", s.Cells["C2"].Value);
                Assert.AreEqual("Joe", s.Cells["C3"].Value);
                Assert.AreEqual(2d, s.Cells["D1"].Value);
                Assert.AreEqual(3d, s.Cells["D2"].Value);
                Assert.AreEqual(1d, s.Cells["D3"].Value);
                Assert.AreEqual("Total", s.Cells["C4"].Value);
                Assert.AreEqual(6d, s.Cells["D4"].Value);
            }
        }

        [TestMethod]
        public void GroupByMixedInput()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["A3"].Value = "NA()";
                s.Cells["A4"].Value = false;
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["B3"].Value = 4;
                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, _xleta.SUM)";
                s.Calculate();
                Assert.AreEqual("Anna", s.Cells["C1"].Value);
                Assert.AreEqual("Joe", s.Cells["C2"].Value);
                Assert.AreEqual(ErrorValues.NAError, s.Cells["C3"].Value);
                Assert.AreEqual(false, s.Cells["C4"].Value);
                Assert.AreEqual(2d, s.Cells["D1"].Value);
                Assert.AreEqual(1d, s.Cells["D2"].Value);
                Assert.AreEqual(4d, s.Cells["D3"].Value);
                Assert.AreEqual(0d, s.Cells["D4"].Value);

                Assert.AreEqual("Total", s.Cells["C5"].Value);
                Assert.AreEqual(7d, s.Cells["D5"].Value);
            }
        }
    }
}
