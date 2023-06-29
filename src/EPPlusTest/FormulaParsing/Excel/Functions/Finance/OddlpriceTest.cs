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
    public class OddlpriceTest
    {
        [TestMethod]
        public void OddlpricedShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,5%,6%,100,2,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 7);
                Assert.AreEqual(97.2372409, result);

            }
        }

        [TestMethod]
        public void OddlpriceIncorrectFrequency()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect frequency: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,1%,5%,100,7,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddlpriceIncorrectRate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect rate: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,-1%,5%,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddlpriceIncorrectPrice()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect yield: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,1%,-1%,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]

        public void OddlpriceIncorrectBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect basis: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,1%,5%,100,2,6)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]

        public void OddlpricedNotGivingBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with no basis (default is zero: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,1%,60%,100,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(36.58083317, result);
            }
        }

        [TestMethod]

        public void OddlpriceIncorrectDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect basis: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 27);
                sheet.Cells["B2"].Value = new System.DateTime(2019, 2, 13);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,1%,5%,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }

        }

        [TestMethod]
        public void OddlpriceExample()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with random inputs: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2029, 2, 22);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLPRICE(B1,B2,B3,19%,529%,1678,4,4)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(29.32151618, result);
            }
        }

    }

}


