using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class LetFunctionTests
    {
        [TestMethod]
        public void LetFunction_SimpleTest()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "LET(x,1 + 2,x + 1)";
            sheet.Calculate();
            Assert.AreEqual(4d, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_WithVariablePrefixes()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "_xlfn.LET(_xlpm.x,B4*B5,_xlpm.y,B6/2,_xlpm.x+_xlpm.y)";
            sheet.Cells["B4"].Value = 4;
            sheet.Cells["B5"].Value = 5;
            sheet.Cells["B6"].Value = 3;
            sheet.Cells["B7"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(21.5, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_WithoutVariablePrefixes()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "LET(x,B4*B5, y,B6/2,x+y)";
            sheet.Cells["B4"].Value = 4;
            sheet.Cells["B5"].Value = 5;
            sheet.Cells["B6"].Value = 3;
            sheet.Cells["B7"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(21.5, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_UsingVariableInOtherVariable()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "_xlfn.LET(_xlpm.x,B4*B5,_xlpm.z, B7, _xlpm.y,B6/_xlpm.z,_xlpm.x+_xlpm.y)";
            sheet.Cells["B4"].Value = 4;
            sheet.Cells["B5"].Value = 5;
            sheet.Cells["B6"].Value = 3;
            sheet.Cells["B7"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(21.5, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_UsingVariableInOtherVariable_WithoutVariablePrefixes()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "LET(x,B4*B5,z, B7, y,B6/z,x+_xlpm.y)";
            sheet.Cells["B4"].Value = 4;
            sheet.Cells["B5"].Value = 5;
            sheet.Cells["B6"].Value = 3;
            sheet.Cells["B7"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(21.5, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_ShouldReturnNameError_RecursiveDeclaration()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "LET(x,1 + x, x + 1)";
            sheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_NestedLets()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A2"].Formula = "LET(x,LET(y,1,y+1),x+1)";
            sheet.Calculate();
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_ShouldReturnAddress1()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["F5"].Formula = "LET(x,B3:B4,x):C4";
            sheet.Cells["B3"].Value = 5;
            sheet.Cells["B4"].Value = 6;
            sheet.Cells["C4"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(5, sheet.Cells["F5"].Value);
            Assert.AreEqual(6, sheet.Cells["F6"].Value);
            Assert.AreEqual(2, sheet.Cells["G6"].Value);
        }

        [TestMethod]
        public void LetFunction_NegateVariable()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(x,3,-x)";
            sheet.Calculate();
            Assert.AreEqual(-3d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LetFunction_UseVariableFromParentLetFunction()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(x,2,y,LET(a,x,a+1),x + y)";
            sheet.Calculate();
            Assert.AreEqual(5d, sheet.Cells["A1"].Value);
        }

        [TestMethod]
        public void LetFunction_ShouldHandleArrayVariableValues()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(x,F1:F2,x + 1)";
            sheet.Cells["F1"].Value = 1;
            sheet.Cells["F2"].Value = 2;
            sheet.Calculate();
            Assert.AreEqual(2d, sheet.Cells["A1"].Value);
            Assert.AreEqual(3d, sheet.Cells["A2"].Value);
        }

        [TestMethod]
        public void LetFunction_ComplexFormula()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Formula = "LET(a,SUM(LET(x,A6,y,A5 + x,x-y),22,LET(x,3,x)),b,LET(x,a-1,x)),c,A3:B4,c * a + b)";
            sheet.Cells["C1"].Value = 43d;
            sheet.Cells["C2"].Value = 65d;
            sheet.Cells["D1"].Value = 241d;
            sheet.Cells["D2"].Value = 263d;
        }
    }
}
