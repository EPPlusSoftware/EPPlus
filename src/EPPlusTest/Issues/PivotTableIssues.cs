using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Xml;
using System.Linq;
using System;
using System.IO;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class PivotTableIssues : TestBase
    {
        [TestMethod]
        public void s688()
        {
            using (ExcelPackage package = OpenTemplatePackage("s688.xlsx"))
            {
                package.Workbook.Worksheets[0].PivotTables[0].Calculate(false);
                SaveAndCleanup(package);
            }
        }
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
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
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
        public void s744()
        {
            using (var p = OpenTemplatePackage("s744.xlsx"))
            {
                ExcelWorkbook workbook = p.Workbook;
                SaveAndCleanup(p);
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
    }
}