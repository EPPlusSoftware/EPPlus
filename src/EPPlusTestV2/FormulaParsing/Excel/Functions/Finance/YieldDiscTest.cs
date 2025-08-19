using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{

    [TestClass]
    public class YieldDiscTest
    {
        [TestMethod]
        public void YieldDiscShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 1;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3,B4,B5)";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(0.053702948, result,0.000000001);
            }
        }

        [TestMethod]
        public void YieldDiscShouldReturnCorrectResult2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 2;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3,B4,B5)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.05282257, result);
            }
        }

        [TestMethod]
        public void YieldDiscShouldReturnCorrectResult3()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3,B4,B5)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.05355622, result);
            }
        }

        [TestMethod]
        public void YieldDiscShouldReturnCorrectResult4()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 0;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3,B4,B5)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.04930107, result);
            }
        }

        [TestMethod]
        public void YieldDiscShouldReturnCorrectResultifEmpty()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3,B4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.04930107, result);
            }
        }
        [TestMethod]
        public void YieldDiscShouldReturnError()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2008, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["C3"].Value = 99.324;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3:C3,B4)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result);
            }
        }
        [TestMethod]
        public void YieldDiscShouldReturnCorrectResult5()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2008, 2, 16);
                sheet.Cells["B2"].Value = new System.DateTime(2018, 3, 1);
                sheet.Cells["B3"].Value = 99.795;
                sheet.Cells["B4"].Value = 100;
                sheet.Cells["B5"].Value = 3;
                sheet.Cells["A1"].Formula = "YIELDDISC(B1,B2,B3,B4,B5)";
                sheet.Calculate();
                var result = (double)sheet.Cells["A1"].Value;
                Assert.AreEqual(0.000204525, result, 0.000000001);
            }
        }
    }
}
