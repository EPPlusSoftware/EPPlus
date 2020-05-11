using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Engineering.Convert
{
    [TestClass]
    public class NumberTestingTests
    {
        [TestMethod]
        public void DeltaTests()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                
                sheet.Cells["A1"].Formula = "DELTA(5, 4)";
                sheet.Calculate();
                Assert.AreEqual(0, sheet.Cells["A1"].Value);

                sheet.Cells["A1"].Formula = "DELTA(1.00001, 1)";
                sheet.Calculate();
                Assert.AreEqual(0, sheet.Cells["A1"].Value);

                sheet.Cells["A1"].Formula = "DELTA(1.23, 1.23)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);

                sheet.Cells["A1"].Formula = "DELTA(1)";
                sheet.Calculate();
                Assert.AreEqual(0, sheet.Cells["A1"].Value);

                sheet.Cells["A1"].Formula = "DELTA(0)";
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value, "DELTA(0) was not 1");
            }
        }
    }
}
