using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ExtLstValidationTests
    {
        [TestMethod]
        public void Issue274()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\Issue274.xlsx")))
            {
                var val = package.Workbook.Worksheets["test1"].DataValidations["A1"];
                Assert.IsInstanceOfType(val, typeof(ExcelDataValidationExtList));
            }
        }

        [TestMethod, Ignore]
        public void AddValidationWithFormulaOnOtherWorksheetShouldReturnExt()
        {
            using(var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("test");
                var sheet2 = package.Workbook.Worksheets.Add("test2");
                var val = sheet1.DataValidations.AddListValidation("A1");
                val.Formula.ExcelFormula = "test2!A1:A2";
                Assert.IsInstanceOfType(val, typeof(ExcelDataValidationExtList));
            }
        }
    }
}
