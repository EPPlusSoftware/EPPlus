using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class MakeArrayTests
    {
        [TestMethod]
        public void MakeArray_SimpleTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(2,2,LAMBDA(r,c,r+c))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
            Assert.AreEqual(3d, sheet.Cells["B1"].Value);
            Assert.AreEqual(4d, sheet.Cells["B2"].Value);
        }
    }
}
