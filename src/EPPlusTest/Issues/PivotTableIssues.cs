﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Xml;
using System.Linq;
using System;
using System.IO;
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

                SaveWorkbook("s692-2.xlsx",p);
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
                SaveWorkbook("i1554-SecondDate.xlsx",package);
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
    }
}