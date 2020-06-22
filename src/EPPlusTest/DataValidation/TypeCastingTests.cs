using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class TypeCastingTests
    {
        [TestMethod]
        public void ShouldCastListValidation()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddListValidation("A1");
                var dv = sheet.DataValidations.First().As.ListValidation;
                Assert.IsNotNull(dv);
            }
        }

        [TestMethod]
        public void ShouldCastIntegerValidation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddIntegerValidation("A1");
                var dv = sheet.DataValidations.First().As.IntegerValidation;
                Assert.IsNotNull(dv);
            }
        }

        [TestMethod]
        public void ShouldCastDecimalValidation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddDecimalValidation("A1");
                var dv = sheet.DataValidations.First().As.DecimalValiation;
                Assert.IsNotNull(dv);
            }
        }

        [TestMethod]
        public void ShouldCastDateTimeValidation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddDateTimeValidation("A1");
                var dv = sheet.DataValidations.First().As.DateTimeValidation;
                Assert.IsNotNull(dv);
            }
        }

        [TestMethod]
        public void ShouldCastTimeValidation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddTimeValidation("A1");
                var dv = sheet.DataValidations.First().As.TimeValidation;
                Assert.IsNotNull(dv);
            }
        }

        [TestMethod]
        public void ShouldCastCustomValidation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddCustomValidation("A1");
                var dv = sheet.DataValidations.First().As.CustomValidation;
                Assert.IsNotNull(dv);
            }
        }

        [TestMethod]
        public void ShouldCastAnyValidation()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.DataValidations.AddAnyValidation("A1");
                var dv = sheet.DataValidations.First().As.AnyValidation;
                Assert.IsNotNull(dv);
            }
        }
    }
}
