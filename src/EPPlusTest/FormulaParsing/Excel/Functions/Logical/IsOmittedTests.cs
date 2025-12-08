using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class IsOmittedTests : TestBase
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
        public void IsOmitted_EmptyInvoke1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(x, IF(ISOMITTED(x), \"saknas\", x))()";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void IsOmitted_Invoke1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(x, IF(ISOMITTED(x), \"saknas\", x))(2)";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void IsOmitted_WithMapUsingRange()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["B1"].Value = 1;
            sheet.Cells["B2"].Value = 2;
            sheet.Cells["A1"].Formula = "MAP(B1:B2, LAMBDA(a, IF(ISOMITTED(a), \"saknas\", a*10)))";
            sheet.Calculate();
            Assert.AreEqual(10d, sheet.Cells["A1"].Value);
            Assert.AreEqual(20d, sheet.Cells["A2"].Value);
        }
    }
}
