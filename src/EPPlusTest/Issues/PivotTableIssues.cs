using Microsoft.CodeCoverage.Core.Reports.Coverage;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class PivotTableIssues : TestBase
    {
        [TestMethod]
        public void s692()
        {
            using (ExcelPackage p = OpenTemplatePackage("s692.xlsx"))
            {
                foreach (ExcelWorksheet worksheet in p.Workbook.Worksheets)
                {
                    foreach (var table in worksheet.PivotTables)
                    {
                        table.Calculate(refreshCache: true);
                    }
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s692_2()
        {
            using (ExcelPackage p = OpenTemplatePackage("s692.xlsx"))
            {
                ExcelWorksheet ws = p.Workbook.Worksheets["data"];

                ws.Cells[2, 1, ws.Dimension.Rows, ws.Dimension.Columns].Clear();
                ws.SetValue(2, 4, "OECD Sustainable consumption behaviour");
                ws.SetValue(2, 9, 1D);
                ws.SetValue(2, 10, 2024D);
                ws.SetValue(2, 11, 4D);
                foreach (ExcelWorksheet worksheet in p.Workbook.Worksheets)
                {
                    foreach (var table in worksheet.PivotTables)
                    {
                        table.Calculate(refreshCache: true);
                    }
                }

                SaveWorkbook("s692-2.xlsx", p);
            }
        }
        [TestMethod]
        public void s713()
        {
            using (ExcelPackage p = OpenTemplatePackage("s713.xlsx"))
            {
                ExcelWorkbook workbook = p.Workbook;
                workbook.Worksheets.Delete("pivot");

                var ns = new XmlNamespaceManager(new NameTable());
                ns.AddNamespace("d", @"http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                var node = workbook.WorkbookXml.SelectSingleNode("//d:pivotCaches", ns);
                if (node != null && node.ChildNodes.Count == 0)
                {
                    node.ParentNode.RemoveChild(node);
                }

                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i1554()
        {
            using (var package = OpenTemplatePackage("i1554.xlsx"))
            {
                AddTableRow(package, 0);
                SaveAndCleanup(package);
            }
            using (var package = OpenPackage("i1554.xlsx"))
            {
                AddTableRow(package, 1);
                var pt = package.Workbook.Worksheets[1].PivotTables[0];
                var cf = pt.Fields[0].Cache;
                cf.Refresh();
                Assert.IsTrue(cf.SharedItems[0] is DateTime);
                Assert.IsTrue(cf.SharedItems[1] is DateTime);
                SaveWorkbook("i1554-SecondDate.xlsx", package);
            }
        }
        private static void AddTableRow(ExcelPackage package, int days)
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets["Data"];
            var table = worksheet.Tables.Single(t => t.Name == "DataTable");
            var column = table.Columns["StartTime"];
            var newRow = table.InsertRow(0);

            newRow.TakeSingleCell(0, column.Position).Value = DateTime.Now.AddDays(days);
            column.DataStyle.NumberFormat.Format = "yyyy-mmmm-dd hh:mm";

            worksheet.Cells[table.Address.Start.Row, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column].AutoFitColumns();
            //workbook.CalculateAllPivotTables(refresh: true);
        }
        [TestMethod]
        public void i1603()
        {
            using (var package = OpenPackage("i1603.xlsx", true))
            {
                var dataSheet = package.Workbook.Worksheets.Add("Data");
                var pivotSheet = package.Workbook.Worksheets.Add("Pivot");

                //put data in the data sheet
                dataSheet.Cells["A1"].Value = "Name";
                dataSheet.Cells["B1"].Value = "Age";
                dataSheet.Cells["C1"].Value = "Gender";

                dataSheet.Cells["A2"].Value = "John";
                dataSheet.Cells["B2"].Value = 25;
                dataSheet.Cells["C2"].Value = "Male";
                dataSheet.Cells["A3"].Value = "Jane";
                dataSheet.Cells["B3"].Value = 30;
                dataSheet.Cells["C3"].Value = "Female";
                dataSheet.Cells["A4"].Value = "Bob";
                dataSheet.Cells["B4"].Value = 40;
                dataSheet.Cells["C4"].Value = "Male";
                dataSheet.Cells["A5"].Value = "Mary";
                dataSheet.Cells["B5"].Value = 28;
                dataSheet.Cells["C5"].Value = "Female";
                dataSheet.Cells["A6"].Value = "John";
                dataSheet.Cells["B6"].Value = 68;
                dataSheet.Cells["C6"].Value = "Male";

                //create pivot table
                var pivotDataRange = dataSheet.Cells[1, 1, 6, 3];
                var pivotTable = pivotSheet.PivotTables.Add(pivotSheet.Cells["C3"], pivotDataRange, "TestPivotTable");

                var field1 = pivotTable.Fields["Name"];
                var f1 = pivotTable.RowFields.Add(field1);
                f1.Items.ShowDetails(false);
                Assert.AreEqual(5, f1.Items.Count);

                var field2 = pivotTable.Fields["Age"];
                var f2 = pivotTable.RowFields.Add(field2);
                f2.Items.ShowDetails(false);
                Assert.AreEqual(6, f2.Items.Count);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s747()
        {
            using (var package = OpenTemplatePackage("s747.xlsx"))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets["Sheet2"];
                worksheet.Cells["A20"].Value = "C";
                worksheet.Cells["A21"].Value = "C";
                worksheet.Cells["A22"].Value = "C";
                worksheet.Cells["A23"].Value = "H";
                worksheet.Cells["A24"].Value = "H";
                worksheet.Cells["A25"].Value = "H";
                worksheet.Cells["B20"].Value = "Test";
                worksheet.Cells["B21"].Value = "Test";
                worksheet.Cells["B22"].Value = "Test";
                worksheet.Cells["B23"].Value = "Test";
                worksheet.Cells["B24"].Value = "Test";
                worksheet.Cells["B25"].Value = "Test";
                worksheet.Cells["C20"].Value = 1;
                worksheet.Cells["C21"].Value = 1;
                worksheet.Cells["C22"].Value = 1;
                worksheet.Cells["C23"].Value = 1;
                worksheet.Cells["C24"].Value = 1;
                worksheet.Cells["C25"].Value = 1;

                var ws2 = workbook.Worksheets["High Level Summary"];
                var pt = ws2.PivotTables[0];
                var slicer1 = ws2.Drawings[0].As.Slicer.PivotTableSlicer;

                Assert.AreEqual(pt.Fields[0].Items.Count, 5);
                Assert.AreEqual(4, slicer1.Cache.Data.Items.Count);
                Assert.AreEqual(false, slicer1.Cache.Data.Items[0].Hidden);
                Assert.AreEqual(false, slicer1.Cache.Data.Items[1].Hidden);
                Assert.AreEqual(true, slicer1.Cache.Data.Items[2].Hidden);
                Assert.AreEqual(true, slicer1.Cache.Data.Items[3].Hidden);

                workbook.CalculateAllPivotTables(true);                              //This causes different but still unexpected changes in the selected values. Happends for true or false

                Assert.AreEqual(6, slicer1.Cache.Data.Items.Count);
                Assert.AreEqual(false, slicer1.Cache.Data.Items[0].Hidden);
                Assert.AreEqual(false, slicer1.Cache.Data.Items[1].Hidden);
                Assert.AreEqual(true, slicer1.Cache.Data.Items[2].Hidden);
                Assert.AreEqual(true, slicer1.Cache.Data.Items[3].Hidden);
                Assert.AreEqual(true, slicer1.Cache.Data.Items[4].Hidden);
                Assert.AreEqual(true, slicer1.Cache.Data.Items[5].Hidden);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void i1713()
        {
            using (var package = OpenTemplatePackage("i1713.xlsx"))
            {
                var dataSheet = package.Workbook.Worksheets["ReportData"];
                var pivotSheet = package.Workbook.Worksheets["Pivot"];
                dataSheet.Calculate();
                //create pivot table
                var pivotDataRange = dataSheet.Cells[3, 1, 28, 20];
                var pivotTable = pivotSheet.PivotTables.Add(pivotSheet.Cells["C3"], pivotDataRange, "TestPivotTable");

                pivotTable.Compact = false;
                (from pf in pivotTable.Fields
                 select pf).ToList().ForEach(f =>
                 {
                     f.Compact = false;
                     f.Outline = false;
                     f.SubtotalTop = false;
                     f.SubTotalFunctions = eSubTotalFunctions.None;
                 });

                //add row fields to pivot table
                var rowField1 = pivotTable.Fields["Group1"];
                pivotTable.RowFields.Add(rowField1);

                var dataField2 = pivotTable.Fields["ID2"];
                var f2 = pivotTable.DataFields.Add(dataField2);
                f2.Name = "Count";
                f2.Function = DataFieldFunctions.Count;

                pivotTable.DataOnRows = false;

                //page field will crush pivot table
                var field = pivotTable.Fields["Data_Missing"];
                var pagef = pivotTable.PageFields.Add(field);

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s744_2()
        {
            using (var p = OpenTemplatePackage("s744-2.xlsx"))
            {
                ExcelWorkbook workbook = p.Workbook;
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void s744_3()
        {
            using (var p = OpenTemplatePackage("FilterClearingExample.xlsx"))
            {
                ExcelWorkbook workbook = p.Workbook;
                p.Workbook.Worksheets[0].PivotTables[0].Calculate();
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void SlicerPivot()
        {
            using (var package = OpenTemplatePackage("Slicer_Empty.xlsx"))
            {
                var wb = package.Workbook;
                foreach (var ws in package.Workbook.Worksheets)
                {
                    foreach (var pTable in ws.PivotTables)
                    {
                        foreach (var field in pTable.Fields)
                        {

                        }
                    }
                }

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void PivotCacheIssue()
        {
            using (var package = OpenTemplatePackage("Issues\\PivotCache\\Sample.xlsx"))
            {
                var wb = package.Workbook;
                foreach (var ws in wb.Worksheets)
                {
                    if (ws.PivotTables.Any()) Console.WriteLine(ws.Name);
                }
                SaveWorkbook("SampleNew.xlsx", package);
            }
        }
        [TestMethod]
        public void PivotErrorCodeIssue()
        {
            using (var p = OpenTemplatePackage("PivotTableIssueErrorCode.xlsx"))
            {
                var wb = p.Workbook;
                foreach (var ws in wb.Worksheets)
                {
                    if (ws.PivotTables.Any())
                        Console.WriteLine(ws.Name);
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void i1968_1()
        {
            using var p = OpenTemplatePackage("i1968-1.xlsx");
            var wb = p.Workbook;

            foreach (var ws in wb.Worksheets)
            {
                if (ws.PivotTables.Any()) Console.WriteLine(ws.Name);
            }

            SaveAndCleanup(p);
        }
        [TestMethod]
        public void i1968_2()
        {
            using var p = OpenTemplatePackage("i1968-2.xlsx");
            var wb = p.Workbook;
            foreach (var ws in wb.Worksheets)
            {
                if (ws.PivotTables.Any()) Console.WriteLine(ws.Name);
            }

            SaveAndCleanup(p);
        }
        [TestMethod]
        public void i1968_2_del()
        {
            using var p = OpenTemplatePackage("i1968-2-del.xlsx");
            var wb = p.Workbook;
            foreach (var ws in wb.Worksheets)
            {
                if (ws.PivotTables.Any()) Console.WriteLine(ws.Name);
            }

            SaveAndCleanup(p);
        }
        public class Shown
        {
            public decimal? Amount { get; set; }
            public DateTime? Date { get; set; }
        }

        [TestMethod]
        public void i820()
        {
            var table = new List<Shown>();
            for (int i = 0; i < 200; i++)
            {
                table.Add(new Shown { Date = DateTime.Today.AddDays(i), Amount = i % 5 == 0 ? 0 : (decimal)10000 });
            }

            using (var pck = OpenPackage("i820.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("data");

                var aa = sheet.Cells["A1"].LoadFromCollection(table, true);
                sheet.Cells["B194"].Value = null;
                sheet.Cells[2, 2, aa.End.Row, 2].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                sheet.Cells[2, 1, aa.End.Row, 1].Style.Numberformat.Format = "#,##0.00";

                var dataRange = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];

                CreatePivotTableWithDataGrouping(pck, dataRange);

                pck.Save();
            }
        }
        [TestMethod]
        public void s880()
        {
            using var p = OpenTemplatePackage("s880.xlsx");
            var ws = p.Workbook.Worksheets["Raw Data"];

            ws.Cells[2, 1].Value = "123456";
            ws.Cells[2, 2].Value = "Doe";
            ws.Cells[2, 3].Value = "John";
            ws.Cells[2, 4].Value = "Example Module";
            ws.Cells[2, 5].Value = Convert.ToDateTime("1/1/2024");
            ws.Cells[2, 6].Value = Convert.ToDateTime("1/1/2025");
            ws.Cells[2, 7].Value = DBNull.Value;
            ws.Cells[2, 8].Value = "Not Registered";
            ws.Cells[2, 9].Value = "Yes";
            //Skip 10 since it's a formula column
            ws.Cells[2, 11].Value = "Example Division";
            ws.Cells[2, 12].Value = "123456 - 111 Main Street";
            ws.Cells[2, 13].Value = "123456 - Job Title";
            ws.Cells[2, 14].Value = DBNull.Value;


            ws.Cells[3, 1].Value = "1234567";
            ws.Cells[3, 2].Value = "Doe";
            ws.Cells[3, 3].Value = "Jane";
            ws.Cells[3, 4].Value = "Example Module";
            ws.Cells[3, 5].Value = Convert.ToDateTime("1/1/2024");
            ws.Cells[3, 6].Value = Convert.ToDateTime("1/1/2025");
            ws.Cells[3, 7].Value = DBNull.Value;
            ws.Cells[3, 8].Value = "Not Registered";
            ws.Cells[3, 9].Value = "Yes";
            //Skip 10 since it's a formula column
            ws.Cells[3, 11].Value = "Example Division";
            ws.Cells[3, 12].Value = "123456 - 111 Main Street";
            ws.Cells[3, 13].Value = "123456 - Job Title";
            ws.Cells[3, 14].Value = "Example Department";


            var headerRow = 1;
            var totalDataRows = 2;
            int firstRowToDelete = totalDataRows + headerRow + 1;
            int deleteCount = ExcelPackage.MaxRows - firstRowToDelete + 1;
            if (deleteCount > 0)
            {
                ws.DeleteRow(firstRowToDelete, deleteCount);
            }

            SaveAndCleanup(p);
        }

        private static void CreatePivotTableWithDataGrouping(ExcelPackage pck, ExcelRangeBase dataRange)
        {
            var wsPivot = pck.Workbook.Worksheets.Add("PivotDateGrp");
            var pt = wsPivot.PivotTables.Add(wsPivot.Cells["B3"], dataRange, "Report");

            //Add a row field
            var rowField = pt.RowFields.Add(pt.Fields["Date"]);
            rowField.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months);

            //Add the data fields and format them
            ExcelPivotTableDataField dataField = pt.DataFields.Add(pt.Fields["Amount"]);
            dataField.Format = "#,##0.00";
            dataField.Name = "Sum of Amount";

            //We want the data fields to appear in columns
            pt.DataOnRows = false;
        }
        [TestMethod]
        public void s877()
        {
            using (var package = OpenTemplatePackage("s877.xlsx"))
            {
                var workbook = package.Workbook;

                var table = new DataTable();
                table.Columns.Add("id", typeof(int));
                table.Columns.Add("Type1", typeof(string));
                table.Columns.Add("Type2", typeof(string));

                table.Rows.Add(4, "c", "z");
                table.Rows.Add(5, "c", "z");
                table.Rows.Add(6, "b", "t");
                table.Rows.Add(7, "b", "t");

                var worksheet = workbook.Worksheets["Sheet1"];
                worksheet.Cells["A5"].LoadFromDataTable(table);

                SaveAndCleanup(package, false);

                //Commenting out this second save results in the expected output
                SaveWorkbook("s877-2.xlsx", package);
            }
        }
        [TestMethod]
        public void s907()
        {
            using (var package = OpenTemplatePackage("s907.xlsx"))
            {
                var firstDataRow = 2;
                var ws = package.Workbook.Worksheets["Raw Data"];


                var firstRowToDelete = firstDataRow + 1;
                ws.DeleteRow(firstRowToDelete, ExcelPackage.MaxRows - firstRowToDelete + 1);
                ws.Cells[firstDataRow, 1].Value = "123456";


                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s907_2()
        {
            using (var package = OpenTemplatePackage("s907-2.xlsx"))
            {
                var firstDataRow = 2;
                var ws = package.Workbook.Worksheets["module_status_report"];


                var firstRowToDelete = firstDataRow + 1;
                ws.DeleteRow(firstRowToDelete, ExcelPackage.MaxRows - firstRowToDelete + 1);
                ws.Cells[firstDataRow, 1].Value = "123456";
                package.GetAsByteArray();

                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public void s910()
        {
            using (var package = OpenTemplatePackage("s910.xlsx"))
            {
                var ws = package.Workbook.Worksheets["Data"];
                var t = ws.Tables[0];

                LoadSomeData(ws, t);

                /*
                 * Call to Columns.Delete below works fine up to EPPlus version 8.0.3
                 * From version 8.0.4 through to 8.0.8 it produces xlsx with incomplete pivot
                 * i.e. "Total in Report Currency" just disappears from the pivot, even though it
                 * is present in the table, and also present in the original pivot definition in the template
                 */

                //delete dummy columns 3..7
                t.Columns.Delete(3, 5);

                Assert.AreEqual(2, package.Workbook.Worksheets["Pivot"].PivotTables[0].DataFields.Count);
                package.Workbook.Calculate();
                SaveWorkbook("s910-Wrong.xlsx", package);
            }
        }
        static void LoadSomeData(ExcelWorksheet ws, OfficeOpenXml.Table.ExcelTable t)
        {
            t.AddRow();
            int row = t.Address.End.Row;
            ws.Cells[row, t.Address.Start.Column + 0].Value = "Client A";
            ws.Cells[row, t.Address.Start.Column + 1].Value = "Jan";
            ws.Cells[row, t.Address.Start.Column + 2].Value = "abc123";
            ws.Cells[row, t.Address.Start.Column + 8].Value = 150.00;

            t.AddRow();
            row = t.Address.End.Row;
            ws.Cells[row, t.Address.Start.Column + 0].Value = "Client A";
            ws.Cells[row, t.Address.Start.Column + 1].Value = "Feb";
            ws.Cells[row, t.Address.Start.Column + 2].Value = "abc22";
            ws.Cells[row, t.Address.Start.Column + 8].Value = 250.00;

            t.AddRow();
            row = t.Address.End.Row;
            ws.Cells[row, t.Address.Start.Column + 0].Value = "Client B";
            ws.Cells[row, t.Address.Start.Column + 1].Value = "Jan";
            ws.Cells[row, t.Address.Start.Column + 2].Value = "cdf43";
            ws.Cells[row, t.Address.Start.Column + 8].Value = 125.00;

            t.DeleteRow(0); // delete 1st/template row
        }
        [TestMethod]
        public void s919()
        {
            using (var p = OpenTemplatePackage("s919.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Aico Data"];
                ws.Calculate();

                Assert.AreEqual(123D ,ws.Cells["C37"].Value);
                Assert.AreEqual(123D, ws.Cells["D38"].Value);
            }
        }
        [TestMethod]
        public void s942()
        {
            using (var p = OpenTemplatePackage("PivotTableCFRemoveTest.xlsx"))
            {
                var ws1 = p.Workbook.Worksheets["Sheet1"];
                var ws2 = p.Workbook.Worksheets["Sheet2"];

                ws1.DeleteRow(2, 5);

                ws2.PivotTables[0].CacheDefinition.Refresh();
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void s993()
        {
            using(var pck = OpenTemplatePackage("s993.xlsx"))
            {
                var calcOpts = new ExcelCalculationOption
                {
                    PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel,
                    AllowCircularReferences = true,
                    EnableUnicodeAwareStringOperations = true
                };

                // Calculate formulas first
                pck.Workbook.Calculate(calcOpts);

                // Refresh + calculate every pivot once
                pck.Workbook.CalculateAllPivotTables(refresh: true);

                // Recalculate any formulas that depend on pivot results
                pck.Workbook.Calculate(calcOpts);

                // Leave workbook in Manual mode and save
                pck.Workbook.CalcMode = ExcelCalcMode.Manual;
                SaveAndCleanup(pck);
            }

            using(var p = OpenPackage("s993.xlsx"))
            {
                var ws = p.Workbook.Worksheets["Data"];

                ws.Calculate();

                // print A1 cell value
                var cell1 = ws.Cells["A1"].Text;

                Assert.AreEqual("60000", cell1);
            }
        }
        [TestMethod]
        public void TestPivot()
        {
            var package = OpenTemplatePackage("TestTemplate3.xlsx");
            var pivotTableCollections = package.Workbook.Worksheets.Select(x => x.PivotTables).ToList();
            // Iterate through each collection of pivot tables and refresh them
            foreach (var pivotTables in pivotTableCollections)
            {
                foreach (var pivotTable in pivotTables)
                {
                    var pivotData = pivotTable.CalculatedData;

                    if (pivotTable.CacheDefinition != null)
                    {
                        pivotTable.CacheDefinition.Refresh();
                        pivotTable.CacheDefinition.SaveData = true;
                    }

                    pivotTable.Calculate(false);
                }
            }
            SaveAndCleanup(package);
        }
        [TestMethod]
        public void TestPivot2()
        {
            var package = OpenTemplatePackage("Bad.xlsx");
            var ws = package.Workbook.Worksheets["Pivottables"];
            var pt = ws.PivotTables[0];
            SaveAndCleanup(package);
        }
        [TestMethod]
        public void s976_1()
        {
            var package = OpenTemplatePackage("s971-1.xlsx");
            var ws = package.Workbook.Worksheets["Monthly Detail"];
            var pt = ws.PivotTables[0];
            Assert.AreEqual(7, ws.ConditionalFormatting.Count);
            Assert.AreEqual(7, pt.ConditionalFormattings.Count);

            //pt.ConditionalFormattings.Clear();
            //Assert.AreEqual(0, ws.ConditionalFormatting.Count);
            //Assert.AreEqual(0, pt.ConditionalFormattings.Count);
            SaveAndCleanup(package);
        }
    }
}