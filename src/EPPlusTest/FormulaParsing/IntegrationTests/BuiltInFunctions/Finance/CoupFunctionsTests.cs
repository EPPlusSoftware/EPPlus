using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.Finance
{
    [TestClass]
    public class CoupFunctionsTests
    {
        [TestMethod]
        public void Couppcd_ShouldReturnCorrectResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COUPPCD(DATE(2017, 05, 30), DATE(2020, 05, 31), 4, 1)";
                sheet.Calculate();
                Assert.AreEqual(new System.DateTime(2017, 2, 28).ToOADate(), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void Coupncd_ShouldReturnCorrectResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COUPNCD(DATE(2017, 02, 01), DATE(2020, 05, 31), 4, 1)";
                sheet.Calculate();
                Assert.AreEqual(new System.DateTime(2017, 2, 28).ToOADate(), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void Coupnum_ShouldReturnCorrectResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COUPNUM(DATE(2016, 02, 01), DATE(2019, 03, 15), 4, 1)";
                sheet.Calculate();
                Assert.AreEqual(13, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void Coupdaysnc_ShouldReturnCorrectResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COUPDAYSNC(DATE(2016, 02, 01), DATE(2019, 05, 31), 4, 1)";
                sheet.Calculate();
                Assert.AreEqual(28d, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void Coupdays_ShouldReturnCorrectResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COUPDAYS(DATE(2012, 2, 29), DATE(2019, 03, 15), 4, 1)";
                sheet.Calculate();
                Assert.AreEqual(91d, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void Coupdaybs_ShouldReturnCorrectResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "COUPDAYBS(DATE(2016, 02, 01), DATE(2019, 05, 31), 2, 1)";
                sheet.Calculate();
                Assert.AreEqual(63, sheet.Cells["A1"].Value);
            }
        }
    }
}
