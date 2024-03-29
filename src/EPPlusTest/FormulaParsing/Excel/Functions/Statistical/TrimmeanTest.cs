﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Statistical
{
    [TestClass]
    public class TrimmeanTest
    {
        [TestMethod]
        public void TrimmeanShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["A6"].Formula = "TRIMMEAN(A1:A5,20%)";
                sheet.Calculate();
                Assert.AreEqual(3d, sheet.Cells["A6"].Value);
            }
        }

        [TestMethod]
        public void TrimmeanShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["A6"].Formula = "TRIMMEAN(A1:A5,110%)";
                sheet.Calculate();
                var result = sheet.Cells["A6"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TrimmeanShouldReturnCorrectResult3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A2"].Value = 2;
                sheet.Cells["A3"].Value = 3;
                sheet.Cells["A4"].Value = 4;
                sheet.Cells["A5"].Value = 5;
                sheet.Cells["A6"].Formula = "TRIMMEAN(A1:A5,-20%)";
                sheet.Calculate();
                var result = sheet.Cells["A6"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void TrimmeanShouldReturnCorrectResult4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["A6"].Formula = "TRIMMEAN(A1,20%)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A6"].Value);
            }
        }
    }
}
