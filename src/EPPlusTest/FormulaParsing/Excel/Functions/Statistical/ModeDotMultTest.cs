using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class ModeMultTests
    {
        [TestMethod]
        public void ModeMult()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["A6"].Formula = "MODE.MULT(A1:A5)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A6"].Value);
                Assert.AreEqual(5d, sheet.Cells["A7"].Value);
            }
        }

        [TestMethod]
        public void ModeMultShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 3;
                sheet.Cells["A6"].Formula = "MODE.MULT(A1:A5)";
                sheet.Calculate();

                Assert.AreEqual(3d, sheet.Cells["A6"].Value);
            }
        }

        [TestMethod]
        public void ModeMultShouldReturnCorrectOrder()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 2;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 4;
                sheet.Cells["A6"].Formula = "MODE.MULT(A1:A5)";
                sheet.Calculate();

                Assert.AreEqual(2d, sheet.Cells["A6"].Value);
                Assert.AreEqual(4d, sheet.Cells["A7"].Value);
            }
        }

        [TestMethod]
        public void ModeMultTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 5;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["A6"].Formula = "MODE.MULT(A1:A5)";
                sheet.Calculate();

                Assert.AreEqual(5d, sheet.Cells["A6"].Value);
            }
        }

        [TestMethod]
        public void ModeMultTest2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 1;
                sheet.Cells["A4"].Value = 2;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["A6"].Formula = "MODE.MULT(A1:A5)";
                sheet.Calculate();

                Assert.AreEqual(1d, sheet.Cells["A6"].Value);
                Assert.AreEqual(2d, sheet.Cells["A7"].Value);
            }
        }


        //[TestMethod]
        //public void ModeMultTest2()
        //{
        //    using (var package = new ExcelPackage())
        //    {
        //        var sheet = package.Workbook.Worksheets.Add("test");
        //        sheet.Cells["A1"].Value = 1;
        //        sheet.Cells["A2"].Value = 2;
        //        sheet.Cells["A3"].Value = 3;
        //        sheet.Cells["A4"].Value = 4;
        //        sheet.Cells["A5"].Value = 5;
        //        sheet.Cells["A6"].Formula = "TRIMMEAN(A1:A5,20%)";
        //        sheet.Calculate();
        //        Assert.AreEqual(3d, sheet.Cells["A6"].Value);
        //    }
        //}
    }
}
