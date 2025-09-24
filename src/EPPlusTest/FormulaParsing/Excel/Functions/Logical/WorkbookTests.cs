using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Logical
{
    [TestClass]
    public class WorkbookTests : TestBase
    {
        [TestMethod]
        public void WorkbookTest1()
        {
            using var p = OpenTemplatePackage("LambdaTests.xlsx");
            Assert.AreEqual(3, p.Workbook.Worksheets.Count);
            p.Workbook.Worksheets["Sheet1"].Calculate();
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void WorkbookTest2()
        {
            using var p = OpenTemplatePackage("LambdaTests.xlsx");
            Assert.AreEqual(3, p.Workbook.Worksheets.Count);
            p.Workbook.Worksheets["Sheet2"].Calculate();
            SaveWorkbook("LambdaTests2.xlsx", p);
        }

        [TestMethod]
        public void WorkbookTest3()
        {
            using var p = OpenTemplatePackage("LambdaTests.xlsx");
            Assert.AreEqual(3, p.Workbook.Worksheets.Count);
            p.Workbook.Worksheets["Sheet3"].Calculate();
            SaveWorkbook("LambdaTests3.xlsx", p);
        }

        [TestMethod]
        public void CreateWorkbookTest()
        {
            using var p = OpenPackage("LambdaWorkbookCreated.xlsx", true);
            var sheet = p.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "MAKEARRAY";
            sheet.Cells["B1"].Formula = "MAKEARRAY(2,2,LAMBDA(r,c,r+c))";
            sheet.Cells["E1"].Value = "SCAN";
            sheet.Cells["F1"].Formula = "SCAN(1,ANCHORARRAY(B1),LAMBDA(a,v,a+1))";
            sheet.Cells["I1"].Value = "MAP";
            sheet.Cells["J1"].Formula = "MAP(ANCHORARRAY(B1),ANCHORARRAY(F1),LAMBDA(a,b,a+b))";
            sheet.Cells["A4"].Value = "BYCOL";
            sheet.Cells["B4"].Formula = "BYCOL(ANCHORARRAY(B1)+1,LAMBDA(array,MAX(array)))";
            sheet.Cells["E4"].Value = "BYROW";
            sheet.Cells["F4"].Formula = "BYROW(ANCHORARRAY(B1),LAMBDA(array,MAX(array)))";
            sheet.Cells["I4"].Value = "ISOMITTED";
            sheet.Cells["J4"].Formula = "ISOMITTED(b)";
            sheet.Calculate();
            SaveAndCleanup(p);
        }
    }
}
