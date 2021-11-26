using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Text;
using System.Globalization;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class TableExporterTests : TestBase
    {
#if !NET35 && !NET40
        [TestMethod]
        public void ShouldExportHeadersAsync()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = "23";
                var table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
                table.TableStyle = TableStyles.Dark1;
                table.ShowHeader = true;
                using(var ms = new MemoryStream())
                {
                    table.HtmlExporter.RenderHtmlAsync(ms).Wait();
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                }
            }
        }
#endif

        [TestMethod]
        public void ShouldExportHeadersWithNoAccessibilityAttributes()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = 23;
                var table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
                table.TableStyle = TableStyles.Dark1;
                table.ShowHeader = true;
                table.HtmlExporter.Settings.Configure(x =>
                { 
                    x.TableId = "myTable"; 
                    x.Minify = true;
                    x.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                });
                var html = table.HtmlExporter.GetHtmlString();
                using (var ms = new MemoryStream())
                {
                    table.HtmlExporter.RenderHtml(ms);
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                    var expectedHtml = "<table class=\"epplus-table ts-dark1 ts-dark1-header ts-dark1-row-stripes\" id=\"myTable\"><thead><tr><th data-datatype=\"string\">Name</th><th data-datatype=\"number\">Age</th></tr></thead><tbody><tr><td>John Doe</td><td data-value=\"23\">23</td></tr></tbody></table>";
                    Assert.AreEqual(expectedHtml, result);
                }
            }
        }
        [TestMethod]
        public void ShouldExportHeadersWithAccessibilityAttributes()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = 23;
                var table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
                table.TableStyle = TableStyles.Dark1;
                table.ShowHeader = true;
                table.HtmlExporter.Settings.Configure(x =>
                {
                    x.TableId = "myTable";
                    x.Minify = true;
                });

                using (var ms = new MemoryStream())
                {
                    table.HtmlExporter.RenderHtml(ms);
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                    var expectedHtml = "<table class=\"epplus-table ts-dark1 ts-dark1-header ts-dark1-row-stripes\" id=\"myTable\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" role=\"columnheader\" scope=\"col\">Name</th><th data-datatype=\"number\" role=\"columnheader\" scope=\"col\">Age</th></tr></thead><tbody role=\"rowgroup\"><tr><td role=\"cell\">John Doe</td><td data-value=\"23\" role=\"cell\">23</td></tr></tbody></table>";
                    Assert.AreEqual(expectedHtml, result);
                }
            }
        }

        [TestMethod]
        public void ExportAllTableStyles()
        {
            string path = _worksheetPath + "TableStyles";
            CreatePathIfNotExists(path);
            using (var p=OpenPackage("TableStylesToHtml.xlsx", true))
            {
                foreach(TableStyles e in Enum.GetValues(typeof(TableStyles)))
                {
                    if (!(e == TableStyles.Custom || e == TableStyles.None))
                    {
                        var ws = p.Workbook.Worksheets.Add(e.ToString());
                        LoadTestdata(ws);
                        var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                        tbl.TableStyle = e;

                        var html = tbl.HtmlExporter.GetSinglePage();

                        File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                    }
                }
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ExportAllFirstLastTableStyles()
        {
            string path = _worksheetPath + "TableStylesFirstLast";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtmlFirstLastCol.xlsx", true))
            {
                foreach (TableStyles e in Enum.GetValues(typeof(TableStyles)))
                {
                    if (!(e == TableStyles.Custom || e == TableStyles.None))
                    {
                        var ws = p.Workbook.Worksheets.Add(e.ToString());
                        LoadTestdata(ws);
                        var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                        tbl.ShowFirstColumn = true;
                        tbl.ShowLastColumn = true;
                        tbl.TableStyle = e;
                        
                        var html = tbl.HtmlExporter.GetSinglePage();

                        File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                    }
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]


        public void ExportAllCustomTableStyles()
        {
            string path = _worksheetPath + "TableStylesCustomFills";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtmlPatternFill.xlsx", true))
            {
                foreach (ExcelFillStyle fs in Enum.GetValues(typeof(ExcelFillStyle)))
                {
                    var ws = p.Workbook.Worksheets.Add($"PatterFill-{fs}");
                    LoadTestdata(ws);
                    var ts = p.Workbook.Styles.CreateTableStyle($"CustomPattern-{fs}", TableStyles.Medium9);
                    ts.FirstRowStripe.Style.Fill.Style = eDxfFillStyle.PatternFill;
                    ts.FirstRowStripe.Style.Fill.PatternType = fs;
                    ts.FirstRowStripe.Style.Fill.PatternColor.Tint=0.10;
                    var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{fs}");
                    tbl.StyleName = ts.Name;

                    var html = tbl.HtmlExporter.GetSinglePage();
                    File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ExportAllGradientTableStyles()
        {
            string path = _worksheetPath + "TableStylesGradientFills";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtmlGradientFill.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add($"PatterFill-Gradient");
                LoadTestdata(ws);
                var ts = p.Workbook.Styles.CreateTableStyle($"CustomPattern-Gradient1", TableStyles.Medium9);
                ts.FirstRowStripe.Style.Fill.Style = eDxfFillStyle.GradientFill;
                ts.FirstRowStripe.Style.Fill.Gradient.GradientType=eDxfGradientFillType.Path;
                var c1 = ts.FirstRowStripe.Style.Fill.Gradient.Colors.Add(0);
                c1.Color.Color = Color.White;

                var c2 = ts.FirstRowStripe.Style.Fill.Gradient.Colors.Add(100);
                c2.Color.Color = Color.FromArgb(0x44, 0x72, 0xc4);
                
                ts.FirstRowStripe.Style.Fill.Gradient.Bottom = 0.5;
                ts.FirstRowStripe.Style.Fill.Gradient.Top = 0.5;
                ts.FirstRowStripe.Style.Fill.Gradient.Left = 0.5;
                ts.FirstRowStripe.Style.Fill.Gradient.Right = 0.5;

                var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tblGradient");
                tbl.StyleName = ts.Name;

                var html = tbl.HtmlExporter.GetSinglePage();
                File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ExportTableWithCellStylesStyles()
        {
            string path = _worksheetPath + "TableStylesCellStyles";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtmlCellStyles.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add($"CellStyles");
                LoadTestdata(ws,100,1,1,true);

                var tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
                tbl.TableStyle = TableStyles.Dark3;
                ws.Cells["A1"].Style.Font.Italic = true;
                ws.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["C5"].Style.Font.Size = 18;
                tbl.Columns[0].TotalsRowLabel = "Total";
                var html = tbl.HtmlExporter.GetSinglePage();
                File.WriteAllText($"{path}\\table-{tbl.StyleName}-CellStyle.html", html);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ShouldExportWithOtherCultureInfo()
        {
            string path = _worksheetPath + "culture";
            CreatePathIfNotExists(path);
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"CellStyles");
                LoadTestdata(ws, 100, 1, 1, true);

                var tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
                tbl.TableStyle = TableStyles.Dark3;
                ws.Cells["A1"].Style.Font.Italic = true;
                ws.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["C5"].Style.Font.Size = 18;
                tbl.Columns[0].TotalsRowLabel = "Total";
                tbl.HtmlExporter.Settings.Culture = new System.Globalization.CultureInfo("en-US");
                
                var html = tbl.HtmlExporter.GetSinglePage();
                File.WriteAllText($"{path}\\table-{tbl.StyleName}-CellStyle.html", html);
            }
        }
        [TestMethod]
        public void ValidateConfigureAndResetToDefault()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"Sheet1");
                LoadTestdata(ws, 100, 1, 1, true);

                var tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
                tbl.TableStyle = TableStyles.Dark3;
                tbl.HtmlExporter.Settings.Configure(x =>
                {
                    x.Encoding = Encoding.Unicode;
                    x.Culture = new CultureInfo("en-GB");
                    x.TableId = "Table1";
                    x.RenderDataAttributes = false;
                    x.Css.Exclude.TableStyle.Border = eBorderExclude.Right | eBorderExclude.Left;
                    x.Css.Exclude.TableStyle.HorizontalAlignment = true;
                    x.Css.Exclude.CellStyle.Fill = true;
                    x.AdditionalTableClassNames.Add("ATC1");
                    x.Accessibility.TableSettings.AriaLabel = "AriaLabel1";
                    x.Accessibility.TableSettings.TableRole = "TableRoll1";
                    x.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                });

                var s = tbl.HtmlExporter.Settings;
                Assert.AreEqual(Encoding.Unicode, s.Encoding);
                Assert.AreEqual("en-GB", s.Culture.Name);
                Assert.AreEqual("Table1", s.TableId);
                Assert.IsFalse(s.RenderDataAttributes);
                Assert.AreEqual(eBorderExclude.Right | eBorderExclude.Left, s.Css.Exclude.TableStyle.Border);
                Assert.IsTrue(s.Css.Exclude.TableStyle.HorizontalAlignment);
                Assert.IsTrue(s.Css.Exclude.CellStyle.Fill);
                Assert.AreEqual("ATC1", s.AdditionalTableClassNames[0]);
                Assert.AreEqual("AriaLabel1", s.Accessibility.TableSettings.AriaLabel);
                Assert.AreEqual("TableRoll1", s.Accessibility.TableSettings.TableRole);
                Assert.IsFalse(s.Accessibility.TableSettings.AddAccessibilityAttributes);

                tbl.HtmlExporter.Settings.ResetToDefault();
                
                s = tbl.HtmlExporter.Settings;
                Assert.AreEqual(Encoding.UTF8, s.Encoding);
                Assert.AreEqual(CultureInfo.CurrentCulture.Name, s.Culture.Name);
                Assert.IsTrue(string.IsNullOrEmpty(s.TableId));
                Assert.IsTrue(s.RenderDataAttributes);
                Assert.AreEqual(0, (int)s.Css.Exclude.TableStyle.Border);
                Assert.IsFalse(s.Css.Exclude.TableStyle.HorizontalAlignment);
                Assert.IsFalse(s.Css.Exclude.CellStyle.Fill);
                Assert.AreEqual(0, s.AdditionalTableClassNames.Count);
                Assert.IsTrue(string.IsNullOrEmpty(s.Accessibility.TableSettings.AriaLabel));
                Assert.AreEqual("table", s.Accessibility.TableSettings.TableRole);
                Assert.IsTrue(s.Accessibility.TableSettings.AddAccessibilityAttributes);
            }
        }

    }
}
