using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class NamedRangeNegationTests
    {
        [TestMethod]
        public void MinusNamedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 123456;

                sheet1.Names.Add("MyRange", sheet1.Cells["A1"]);

                sheet1.Cells["C3"].Formula = "-MyRange";

                pck.Workbook.Calculate();

                Assert.AreEqual(-123456, sheet1.Cells["C3"].GetValue<double>(), 1E-5); //ERROR: evaluates to 123456
            }
        }
        [TestMethod]
        public void MinusNamedRangePlusNamedRange()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet1 = pck.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["A1"].Value = 123456;
                sheet1.Cells["B1"].Value = 3;

                sheet1.Names.Add("MyRange", sheet1.Cells["A1"]);
                sheet1.Names.Add("Another", sheet1.Cells["B1"]);

                sheet1.Cells["C3"].Formula = "-MyRange+Another";

                pck.Workbook.Calculate();

                Assert.AreEqual(-123453, sheet1.Cells["C3"].GetValue<double>(), 1E-5); //ERROR: evaluates to 123459
            }
        }
    }
}
