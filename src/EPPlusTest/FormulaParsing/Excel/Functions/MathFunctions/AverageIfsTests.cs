using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.MathFunctions
{
    [TestClass]
    public class AverageIfsTests
    {
        [TestMethod]
        public void AverageIfsShouldNotCountNumericStringsAsNumbers()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[2, 1].Value = 4;
                sheet.Cells[3, 1].Value = 5;
                sheet.Cells[1, 2].Value = 1;
                sheet.Cells[2, 2].Value = "2";
                sheet.Cells[3, 2].Value = 3;
                sheet.Cells[1, 3].Value = 2;
                sheet.Cells[2, 3].Value = 1;
                sheet.Cells[3, 3].Value = "4";

                sheet.Cells[4, 1].Formula = "AVERAGEIFS(A1:A3,B1:B3,\">0\",C1:C3,\">1\")";
                sheet.Calculate();
                var val = sheet.Cells[4, 1].Value;
                Assert.AreEqual(3d, val);
            }
        }
        [TestMethod]
        public void ShouldHandleErrorInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells[1, 1].Value = 3;
                sheet.Cells[2, 1].Value = 4;
                sheet.Cells[3, 1].Value = 5;
                sheet.Cells[1, 2].Value = "#REF!";
                sheet.Cells[2, 2].Value = new ExcelErrorValue(eErrorType.Ref); 
                sheet.Cells[3, 2].Value = 3;

                sheet.Cells[4, 1].Formula = "AVERAGEIFS(A1:A3,B1:B3, #REF!)";
                sheet.Calculate();
                var val = sheet.Cells[4, 1].Value;
                Assert.AreEqual(4d, val);
            }
        }

        [TestMethod]
        public void AverageIfsShouldIgnoreErrorsInRangeIfInCriteria()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 2;
                sheet.Cells["C1"].Value = 3;
                sheet.Cells["A2"].Value = "a";
                sheet.Cells["B2"].Value = ErrorValues.NAError;
                sheet.Cells["C2"].Value = "Test";

                sheet.Cells["A3"].Formula = "AVERAGEIFS(A1:C1,A2:C2,\"=#N/A\")";
                sheet.Calculate();
                Assert.AreEqual(2d, sheet.Cells["A3"].Value);

                sheet.Cells["A3"].Formula = "AVERAGEIFS(A1:C1,A2:C2,\"=a\")";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["A3"].Value);
            }
        }
    }
}
