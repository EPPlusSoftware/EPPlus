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
        public void LetTest1()
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
        public void LetTest_UsingVariableInOtherVariable()
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
    }
}
