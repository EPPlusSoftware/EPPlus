using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Xml;
using System.Linq;
using System;
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
    }
}
