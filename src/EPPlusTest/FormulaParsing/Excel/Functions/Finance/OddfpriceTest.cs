using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class OddfpriceTest : TestBase
    {
        [TestMethod]
        public void OddfpriceShortPeriodTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with short period");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 01);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 12, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                //sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,5%,6%,100,2,0)";
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,2,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                //Assert.AreEqual(97.26007079, result);
                Assert.AreEqual(86.29690031, result);


            }
        }

        [TestMethod]
        public void OddfpriceLongPeriodTest()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with long period:");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,4,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(86.35406228, result);
            }
        }

        [TestMethod]
        public void OddfpriceInvalidBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with invalid basis: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,4,8)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfpriceIncorrectRate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with negative rate: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,-5%,6%,100,4,8)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfpriceIncorrectYield()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with negative yield: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,-6%,100,4,1)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }

        }

        [TestMethod]
        public void OddfpriceInvalidDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test should return error when dates are incorrect: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2017, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,5%,100,4,1)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfpriceWithLowFrequency()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with coupon frequency of 1 year periods:");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,1,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(86.62364985, result);
            }

        }

        [TestMethod]
        public void OddfpriceWithIncorrectFrequency()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect frequency: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,5,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddfpriceNoBasisArgument()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with no basis argument: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 02, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 02, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 01);
                sheet.Cells["B4"].Value = new System.DateTime(2019, 02, 15);
                sheet.Cells["A1"].Formula = "ODDFPRICE(B1,B2,B3,B4,1%,6%,100,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(86.62364985, result);
            }
        }
    }

}