using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class SumIfsTests
    {
        [TestMethod, Ignore]
        public void SumIfsShouldHandleSingleRange()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "SUMIFS(H5,H5,\">0\",K5,\"> 0\")";
                sheet.Cells["H5"].Value = 1;
                sheet.Cells["K5"].Value = 1;
                sheet.Calculate();
                Assert.AreEqual(1, sheet.Cells["A1"].Value);
            }
        }
    }
}
