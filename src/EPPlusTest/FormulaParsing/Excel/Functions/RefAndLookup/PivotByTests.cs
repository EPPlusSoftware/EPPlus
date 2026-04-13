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
        public void PivotBy()
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
            }
        }
    }
}
