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
    }
}
