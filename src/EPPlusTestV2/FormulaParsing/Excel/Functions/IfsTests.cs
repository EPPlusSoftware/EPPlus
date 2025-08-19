using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class IfsTests
    {
        [TestMethod]
        public void IfsShouldReturnFirstTrueArg()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(1>2,1,2>1,2)";
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldHaveAtLeastTwoArgs()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(1>2,)";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Parse("#N/A"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldReturnValueErrorIfUnevenNumberOfArgs()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(1>2,1,2>1)";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Parse("#VALUE!"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldReturnNaErrorIfNoArgsEvaluatesToTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(1>2,1,1>1,2)";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldReturnValueErrorIfNoArgsEvaluatesToTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(1>2,1,1>1,2)";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldHandleNonZeroNumericValueAs0()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(2,1,1>1,2)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldHandleDivisionByZero()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(0,1,1,1/0)";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldIgnoreDivisionByZeroAfterTrueCondition()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(1,1,1,1/0)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldValueErrorIfNonBoolOrNumericArg()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "IFS(\"A\",1,1,1/0)";
                sheet.Calculate();
                Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void IfsShouldHandleRange()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["B2"].Value = 1;
                sheet.Cells["C1"].Value = 2;
                sheet.Cells["C2"].Value = 4;
                sheet.Cells["A1"].Formula = "IFS(B1>C1,3,B2<C2,4)";
                sheet.Calculate();
                Assert.AreEqual(4d, sheet.Cells["A1"].Value);
            }
        }
    }
}
