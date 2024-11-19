using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;
using System;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class TableIssues : TestBase
    {
        [TestMethod]
        public void s594()
        {
            using (ExcelPackage package = OpenTemplatePackage("s594.xlsx"))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["dg"];

                ExcelCalculationOption excelCalculationOption = new ExcelCalculationOption();
                excelCalculationOption.AllowCircularReferences = true;
                worksheet.Calculate(excelCalculationOption);

                Assert.AreNotEqual(0, worksheet.Cells["A1"].Text);

                package.Save();
            }
        }
        [TestMethod]
        public void i1314()
        {
            using (var p = OpenTemplatePackage("i1314.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var tbl = ws.Tables[0];
                tbl.InsertRow(1,1);
				tbl.AddRow(1);

				SaveAndCleanup(p);
            }
		}
        [TestMethod]
        public void i1642()
        {
            using (var package = OpenTemplatePackage("i1642.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                var excelTable = worksheet.Tables[0];
                
                var col = excelTable.Range.Offset(0, 10).TakeSingleColumn(0).SkipRows(1);
                var formulaStr = col.TakeSingleCell(0, 0).Formula;
                col.CreateArrayFormula(formulaStr, true);
                SaveAndCleanup(package);
            }
        }
    }
}
