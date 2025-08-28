using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class MakeArrayTests
    {
        [TestMethod]
        public void MakeArray_SimpleTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(2,2,LAMBDA(r,c,r+c))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
            Assert.AreEqual(3d, sheet.Cells["B1"].Value);
            Assert.AreEqual(4d, sheet.Cells["B2"].Value);
        }

        [TestMethod]
        public void MakeArray_Test2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["D4"].Formula = "MAKEARRAY(D2,E2,LAMBDA(row,col,CHOOSE(RANDBETWEEN(1,3),\"Red\",\"Blue\",\"Green\")))";
            sheet.Cells["D2"].Value = 10;
            sheet.Cells["E2"].Value = 1;
            sheet.Calculate();
            bool IsValidCellValue(object val)
            {
                var arr = new string[] { "Red", "Blue", "Green" };
                foreach (var str in arr)
                {
                    if (string.Compare(str, val.ToString()) == 0) return true;
                }
                return false;
            };
            var d4 = sheet.Cells["D4"].Value;
            var d5 = sheet.Cells["D5"].Value;
            Assert.IsTrue(IsValidCellValue(d4));
            Assert.IsTrue(IsValidCellValue(d5));
        }

        [TestMethod]
        public void MakeArray_Test3()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(LAMBDA(a,b,a+b)(1,2),3,LAMBDA(r,c,r+c))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
            Assert.AreEqual(4d, sheet.Cells["A3"].Value);
            Assert.AreEqual(3d, sheet.Cells["B1"].Value);
            Assert.AreEqual(4d, sheet.Cells["B2"].Value);
            Assert.AreEqual(5d, sheet.Cells["B3"].Value);
            Assert.AreEqual(4d, sheet.Cells["C1"].Value);
            Assert.AreEqual(5d, sheet.Cells["C2"].Value);
            Assert.AreEqual(6d, sheet.Cells["C3"].Value);
        }

        [TestMethod]
        public void MakeArray_Test4()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(LAMBDA(a,b,a+b)(1,2),LAMBDA(x, x * 1)(LAMBDA(a, a)(3)),LAMBDA(r,c,r+c))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
            Assert.AreEqual(4d, sheet.Cells["A3"].Value);
            Assert.AreEqual(3d, sheet.Cells["B1"].Value);
            Assert.AreEqual(4d, sheet.Cells["B2"].Value);
            Assert.AreEqual(5d, sheet.Cells["B3"].Value);
            Assert.AreEqual(4d, sheet.Cells["C1"].Value);
            Assert.AreEqual(5d, sheet.Cells["C2"].Value);
            Assert.AreEqual(6d, sheet.Cells["C3"].Value);
        }


        [TestMethod]
        public void MakeArray_Test5()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(3, LAMBDA(x, x)(LAMBDA(a, a)(3)), LAMBDA(r,c, r+c))";
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
            Assert.AreEqual(6d, sheet.Cells["C3"].Value);
        }

        [TestMethod]
        public void TokenizerTest()
        {
            var str = "MAKEARRAY(3, LAMBDA(x, x)(LAMBDA(a, a)(3)), LAMBDA(r,c, r+c))";
            var tokens = SourceCodeTokenizer.Default.Tokenize(str);
            var t = tokens[9];
            var rpn = FormulaExecutor.CreateRPNTokens(tokens);
        }

        [TestMethod]
        public void MakeArray_Test6()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LAMBDA(x,x)(LAMBDA(y,y)(1))";
            sheet.Calculate();
            Assert.AreEqual(1d, sheet.Cells["A1"].Value);
        }


        [TestMethod]
        public void MakeArray_ShouldReturnValueErrorIfWrongNumberOfArgsInLambda1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(LAMBDA(a,b,a+b)(1,2),3,LAMBDA(r,c,z,r+c+z))";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void MakeArray_ShouldReturnValueErrorIfWrongNumberOfArgsInLambda2()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "MAKEARRAY(LAMBDA(a,b,a+b)(1,2),3,LAMBDA(r,r+1)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void MakeArray_WithOuterLet()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["C1"].Formula = "LET(n, 3, MAKEARRAY(n, n, LAMBDA(r,c,r*c)))";
            sheet.Calculate();
            Assert.AreEqual(1d, sheet.Cells["C1"].Value); // r=1, c=1
            Assert.AreEqual(4d, sheet.Cells["D2"].Value); // r=2, c=2
            Assert.AreEqual(9d, sheet.Cells["E3"].Value); // r=3, c=3
        }
    }
}
