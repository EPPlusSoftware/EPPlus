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
    public class OddlyieldTest
    {
        [TestMethod]
        public void OddlyieldShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with correct input values");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,1%,34,100,1,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.67023199, result);

            }
        }

        [TestMethod]
        public void OddlyieldIncorrectFrequency()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect frequency: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,1%,34,100,7,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddlyieldIncorrectRate()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect rate: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,-1%,34,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]
        public void OddlyieldIncorrectPrice()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect price: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,1%,-34,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]

        public void OddlyieldIncorrectBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect basis: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,1%,34,100,2,6)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }
        }

        [TestMethod]

        public void OddlyieldNotGivingBasis()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with no basis (default is zero: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2022, 2, 15);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,1%,34,100,1)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A1"].Value, 8);
                Assert.AreEqual(0.67023199, result);
            }
        }

        [TestMethod]

        public void OddlyieldIncorrectDates()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with incorrect dates: ");
                sheet.Cells["B1"].Value = new System.DateTime(2019, 2, 14);
                sheet.Cells["B2"].Value = new System.DateTime(2019, 2, 13);
                sheet.Cells["B3"].Value = new System.DateTime(2018, 11, 1);
                sheet.Cells["A1"].Formula = "ODDLYIELD(B1,B2,B3,1%,34,100,2,0)";
                sheet.Calculate();
                var result = sheet.Cells["A1"].Value;
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), result);
            }

        }

        [TestMethod]

        public void OddlyieldExample()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test with random example: ");
                sheet.Cells["A1"].Value = new System.DateTime(2018, 2, 05);
                sheet.Cells["A2"].Value = new System.DateTime(2018, 06, 15);
                sheet.Cells["A3"].Value = new System.DateTime(2017, 10, 15);
                sheet.Cells["A4"].Formula = "ODDLYIELD(A1,A2,A3,5%,99.5,100,2,0)";
                sheet.Calculate();
                var result = System.Math.Round((double)sheet.Cells["A4"].Value, 9);
                Assert.AreEqual(0.063196633, result);
            }
        }

    }
}
