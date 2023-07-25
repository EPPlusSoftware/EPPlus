using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class LognormDotInvTests
    {
        [TestMethod]
        public void LognormDotInvTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A2"].Value = 0.039084;
                sheet.Cells["A3"].Value = 3.5;
                sheet.Cells["A4"].Value = 1.2;

                sheet.Cells["B2"].Formula = "LOGNORM.INV(A2,A3,A4)";
                sheet.Calculate();

                var result = System.Math.Round((double)sheet.Cells["B2"].Value, 7);
                Assert.AreEqual(4.0000252, result);
            }
        }
    }
}