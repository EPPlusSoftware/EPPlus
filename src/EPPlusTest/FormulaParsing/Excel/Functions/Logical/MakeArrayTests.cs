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

        [TestMethod]
        public void MakeArray_Test2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["D4"].Formula = "MAKEARRAY(D2,E2,LAMBDA(row,col,CHOOSE(RANDBETWEEN(1,3),\"Red\",\"Blue\",\"Green\")))";
            sheet.Cells["D2"].Value = 10;
            sheet.Cells["E2"].Value = 1;
            sheet.Calculate();
            bool IsValidCellValue(object val)
            {
                var arr = new string[] { "Red", "Blue", "Green" };
                foreach (var str in arr)
                {
                    if (string.Compare(str, val.ToString()) == 0) return true;
                }
                return false;
            };
            var d4 = sheet.Cells["D4"].Value;
            var d5 = sheet.Cells["D5"].Value;
            Assert.IsTrue(IsValidCellValue(d4));
            Assert.IsTrue(IsValidCellValue(d5));
        }

        [TestMethod]
        public void MakeArray_Test3()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(LAMBDA(a,b,a+b)(1,2),3,LAMBDA(r,c,r+c))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
            Assert.AreEqual(4d, sheet.Cells["A3"].Value);
            Assert.AreEqual(3d, sheet.Cells["B1"].Value);
            Assert.AreEqual(4d, sheet.Cells["B2"].Value);
            Assert.AreEqual(5d, sheet.Cells["B3"].Value);
            Assert.AreEqual(4d, sheet.Cells["C1"].Value);
            Assert.AreEqual(5d, sheet.Cells["C2"].Value);
            Assert.AreEqual(6d, sheet.Cells["C3"].Value);
        }
    }
}
