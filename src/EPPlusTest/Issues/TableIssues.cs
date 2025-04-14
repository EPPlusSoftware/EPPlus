using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Data;

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

                Assert.AreNotEqual("0", worksheet.Cells["A1"].Text);

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

        [TestMethod]
        public void sc813()
        {

            var dataTable = new DataTable();

            dataTable.Columns.Add("A", typeof(string));
            dataTable.Columns.Add("B", typeof(string));
            dataTable.Columns.Add("C", typeof(string));
            dataTable.Columns.Add("D", typeof(string));
            dataTable.Columns.Add("E", typeof(string));

            using var package = OpenPackage("sc813.xlsx", true);
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            var range = worksheet.Cells["A2"].LoadFromDataTable(dataTable, true);
            var table = worksheet.Tables.Add(range, "TestTable");
            table.ShowHeader = true;

            //Initial issue: Commenting either of these insert/load combos will result in a corrupted workbook
            table.InsertRow(int.MaxValue, 5);
            worksheet.Cells[table.Address.End.Row, table.Address.Start.Column].LoadFromArrays(new List<object[]> { new[] { "1", "2", "3", "4", "5" } });
            table.InsertRow(int.MaxValue, 5);
            worksheet.Cells[table.Address.End.Row, table.Address.Start.Column].LoadFromArrays(new List<object[]> { new[] { "z", "x", "y", "x", "w" } });


            SaveAndCleanup(package);
        }

        /// <summary>
        /// i1885
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void EpplusShouldThrowOnMultiCellArrayFormulaInTable()
        {
            using (var package = OpenPackage("Multi-CellArrayFormulaInTable.xlsx", true))
            {
                var wb = package.Workbook;
                var sheet = wb.Worksheets.Add("newWorksheet");

                var excelTable = sheet.Tables.Add(sheet.Cells["A1:D4"], "TableTest");

                sheet.Cells["D2:D3"].CreateArrayFormula("SUM(A2:B2 * 1)", true);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void IntersectTableShouldWork()
        {
            using (var package = OpenPackage("TableIntersect.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("newWorksheet");

                var excelTable = ws.Tables.Add(ws.Cells["E3:H10"], "TableTest");

                var intersectsCovered = ws.Cells["F5"].IntersectsWithTable();
                var intersectsLeft = ws.Cells["D4:E4"].IntersectsWithTable();
                var intersectsTop = ws.Cells["E2:E3"].IntersectsWithTable();

                var noIntersectTop = ws.Cells["E2:F2"].IntersectsWithTable();
                var noIntersectBot = ws.Cells["E11:E12"].IntersectsWithTable();

                Assert.IsTrue(intersectsCovered);
                Assert.IsTrue(intersectsLeft);
                Assert.IsTrue(intersectsTop);

                Assert.IsFalse(noIntersectTop);
                Assert.IsFalse(noIntersectBot);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void IntersectTableShouldWorkInsertDelete()
        {
            using (var package = OpenPackage("TableIntersectInsertDelete.xlsx", true))
            {
                var wb = package.Workbook;
                var ws = wb.Worksheets.Add("newWorksheet");

                var excelTable = ws.Tables.Add(ws.Cells["E3:H10"], "TableTest");

                ws.InsertColumn(2, 1);

                //False after insert
                var intersectsLeft = ws.Cells["D4:E4"].IntersectsWithTable();
                Assert.IsFalse(intersectsLeft);

                //True after insert
                var intersectsRight = ws.Cells["I4:I4"].IntersectsWithTable();
                Assert.IsTrue(intersectsRight);

                ws.DeleteColumn(2);

                //True after delete
                intersectsLeft = ws.Cells["D4:E4"].IntersectsWithTable();
                Assert.IsTrue(intersectsRight);

                //False after delete
                intersectsRight = ws.Cells["I4:I4"].IntersectsWithTable();
                Assert.IsFalse(intersectsRight);

                SaveAndCleanup(package);
            }
        }
    }
}
