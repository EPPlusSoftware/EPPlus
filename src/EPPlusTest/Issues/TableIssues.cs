using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeOpenXml.Table;

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

        /// <summary>
        /// Same as i1642 but in english
        /// </summary>
        [TestMethod]
        public void TableFormulaTest()
        {
            using (var package = OpenTemplatePackage("tableArrayTest.xlsx"))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                var excelTable = worksheet.Tables[0];

                var aForm = worksheet.Cells["G2"].FormulaR1C1;
                worksheet.Cells["G2:G5"].FormulaR1C1 = aForm;

                worksheet.Calculate();

                Assert.AreEqual(2d, worksheet.Cells["G2"].Value);
                Assert.AreEqual(4d, worksheet.Cells["G3"].Value);
                Assert.AreEqual(6d, worksheet.Cells["G4"].Value);
                Assert.AreEqual(8d, worksheet.Cells["G5"].Value);

                Assert.AreEqual("Table1[[#This Row],[Column3]]+M2", worksheet.Cells["G2"].Formula);
                Assert.AreEqual("Table1[[#This Row],[Column3]]+M3", worksheet.Cells["G3"].Formula);
                Assert.AreEqual("Table1[[#This Row],[Column3]]+M4", worksheet.Cells["G4"].Formula);
                Assert.AreEqual("Table1[[#This Row],[Column3]]+M5", worksheet.Cells["G5"].Formula);

                SaveAndCleanup(package);
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
                col.ClearFormulaValues();
                col.ClearFormulas();
                col.Formula = formulaStr;

                worksheet.Calculate();

                Assert.AreEqual(2d, worksheet.Cells["K2"].Value);
                Assert.AreEqual(4d, worksheet.Cells["K3"].Value);
                Assert.AreEqual(6d, worksheet.Cells["K4"].Value);
                Assert.AreEqual(8d, worksheet.Cells["K5"].Value);

                Assert.AreEqual("表1[[#This Row],[列5]]+M2", worksheet.Cells["K2"].Formula);
                Assert.AreEqual("表1[[#This Row],[列5]]+M3", worksheet.Cells["K3"].Formula);
                Assert.AreEqual("表1[[#This Row],[列5]]+M4", worksheet.Cells["K4"].Formula);
                Assert.AreEqual("表1[[#This Row],[列5]]+M5", worksheet.Cells["K5"].Formula);

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
        [TestMethod]
        public void StructuredReferenceShouldWorkAsExpected()
        {
            using (var package = OpenPackage("StructuredReference.xlsx", true))
            {
                package.Workbook.FullCalcOnLoad = false;

                var ws = package.Workbook.Worksheets.Add("name");

                var aTable = ws.Tables.Add(ws.Cells["A1:D10"], "ATable");
                ws.Cells["A2:D10"].Formula = "ROW()+COLUMN()";
                aTable.ShowHeader = true;

                ws.Cells["D1"].Value = "Space Separated";

                ws.Calculate();

                aTable.SyncColumnNames(ApplyDataFrom.CellsToColumnNames);

                ws.Cells["G1"].Formula = "ATable[#Headers]";

                ws.Cells["A20"].Formula = "ATable[#Data]";

                ws.Cells["H5"].Formula = "ATable[Column2]";

                ws.Cells["I20"].Formula = "ATable[[#Headers],[#Data],[Column3]]";

                ws.Cells["P1"].Formula = "ATable[#All]";

                ws.Cells["P20"].Formula = "ATable[[#Headers],[#Data],[Column2]:[Column3]]";

                ws.Cells["Z1"].Formula = "ATable[Space Separated]";

                ws.Calculate();

                var headerResRange = ws.Cells["G1:J1"];
                var headerResValues = headerResRange.Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                var colNamesList = aTable.Columns.GetColNamesList();
                Assert.IsTrue(headerResValues.SequenceEqual(colNamesList));

                var origStrings = ws.Cells["A2:D10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());
                var resultString = ws.Cells["A20:D28"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                Assert.IsTrue(origStrings.SequenceEqual(resultString));

                var origStringsCol2 = ws.Cells["B2:B10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());
                var resultStringsCol2 = ws.Cells["H5:H13"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                Assert.IsTrue(origStringsCol2.SequenceEqual(resultStringsCol2));

                var origStringsCol3 = ws.Cells["C1:C10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());
                var resultStringsCol3 = ws.Cells["I20:I29"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                Assert.IsTrue(origStringsCol3.SequenceEqual(resultStringsCol3));

                var origStringsAll = ws.Cells["A1:D10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());
                var resultStringsAll = ws.Cells["P1:S10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                Assert.IsTrue(origStringsAll.SequenceEqual(resultStringsAll));

                var origStrings2Columns = ws.Cells["B1:C10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());
                var resultStrings2Columns = ws.Cells["P20:Q29"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                Assert.IsTrue(origStrings2Columns.SequenceEqual(resultStrings2Columns));

                var origStringsSpaceSep = ws.Cells["D2:D10"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());
                var resultStrings2SpaceSep = ws.Cells["Z1:Z9"].Where(x => x.Value != null).Select(y => y.GetCellValue<string>());

                Assert.IsTrue(origStringsSpaceSep.SequenceEqual(resultStrings2SpaceSep));

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i2081_EpplusGenerated()
        {
            using (var p = OpenPackage("i2081_Generated.xlsx", true))
            {
                //Set up the workbook as example
                using var workbook = p.Workbook;
                var wsName = "singleCellTable";
                var ws = workbook.Worksheets.Add(wsName);

                var tableRange = ws.Cells["A1:A2"];

                var scTable = ws.Tables.Add(tableRange, "Table1");

                scTable.ShowHeader = true;
                scTable.ShowTotal = true;
                scTable.Columns[0].TotalsRowFunction = RowFunctions.Sum;

                tableRange.AutoFitColumns();
                ws.Calculate();

                //Perform test
                var values = new object[,] { { 123 } };
                ws.Cells["A2:A2"].Value = values;

                var logfile = new FileInfo("epplus_i2081_Log.txt");

                workbook.FormulaParserManager.AttachLogger(logfile);

                ws.Calculate();

                workbook.FormulaParserManager.DetachLogger();

                var sr = logfile.OpenText();
                var logStr = sr.ReadToEnd();

                sr.Close();
                logfile.Delete();

                Assert.IsTrue(logStr.Contains($"Set value in Cell\t{wsName}!A3\t123\tDecimal"));
                Assert.AreEqual(values[0, 0], ws.Cells["A2:A2"].Value);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i2081_SingleMultArr()
        {
            using (var p = OpenPackage("i2081_SingleCellMultArr.xlsx", true))
            {
                using var workbook = p.Workbook;
                var wsName = "singleCellTable";
                var ws = workbook.Worksheets.Add(wsName);

                var tableRange = ws.Cells["A1:A2"];

                var scTable = ws.Tables.Add(tableRange, "Table1");

                scTable.ShowHeader = true;
                scTable.ShowTotal = true;
                scTable.Columns[0].TotalsRowFunction = RowFunctions.Sum;

                tableRange.AutoFitColumns();
                ws.Calculate();
                //Part that is different START
                var values = new object[,] { { 1, 123 }, { 2, 456 } };
                ws.Cells["A2:A2"].Value = values;
                //Part that is different END

                var logfile = new FileInfo("epplus_i2081MultArr_Log.txt");

                workbook.FormulaParserManager.AttachLogger(logfile);

                ws.Calculate();

                workbook.FormulaParserManager.DetachLogger();

                var sr = logfile.OpenText();
                var logStr = sr.ReadToEnd();

                sr.Close();
                logfile.Delete();

                Assert.IsTrue(logStr.Contains($"Set value in Cell\t{wsName}!A3\t1\tDecimal"));
                Assert.AreEqual(values[0, 0], ws.Cells["A2:A2"].Value);

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i2081_Formula()
        {
            using (var p = OpenPackage("i2081_Formulas.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Aws");
                ws.Cells["C3"].Formula = "{1,2,3;4,5,6}";

                ws.Calculate();

                Assert.AreEqual(1, ws.Cells["C3"].Value);
                Assert.AreEqual(2, ws.Cells["D3"].Value);
                Assert.AreEqual(3, ws.Cells["E3"].Value);

                Assert.AreEqual(4, ws.Cells["C4"].Value);
                Assert.AreEqual(5, ws.Cells["D4"].Value);
                Assert.AreEqual(6, ws.Cells["E4"].Value);



                SaveAndCleanup(p);
            }
        }
    }
}
