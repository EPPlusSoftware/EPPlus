using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class IsOmittedTests
    {
        [TestMethod]
        public void IsOmitted_ShouldReturnFalseWhenOutOfLambdaScope()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "ISOMITTED(b)";
            sheet.Calculate();
            bool? result = null;
            if(sheet.Cells["A1"].Value is bool b)
            {
                result = b;
            }
            Assert.IsNotNull(result);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void IsOmitted_ShouldReturnTrueIfParamIsNull()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1,) + 1";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void IsOmitted_ShouldReturnFalseIfParamIsNull()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(a,b,IF(ISOMITTED(b),a,a+b))(1,2) + 1";
            sheet.Calculate();
            Assert.AreEqual(4d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void IsOmitted_EmptyInvoke()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(x, IF(ISOMITTED(x), \"saknas\", x))()";
            sheet.Calculate();
            Assert.AreEqual("saknas", sheet.Cells["A1"].Value);
        }
    }
}
