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
using System.Threading.Tasks;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class TableExporterTests : TestBase
    {
        string _htmlOutput;
        public TableExporterTests() : base()
        {
            _htmlOutput = _worksheetPath + "\\html\\";
            if (Directory.Exists(_htmlOutput) == false)
            {
                Directory.CreateDirectory(_htmlOutput);
            }
        }
#if !NET35 && !NET40
        [TestMethod]
        public void ShouldExportHeadersAsync()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = "23";
                var table = sheet.Tables.Add(sheet.Cells["A1:B2"], "myTable");
                table.TableStyle = TableStyles.Dark1;
                table.ShowHeader = true;
                using (var ms = new MemoryStream())
                {
                    var exporter = table.CreateHtmlExporter();
                    exporter.RenderHtmlAsync(ms).Wait();
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
                var exporter = table.CreateHtmlExporter();
                exporter.Settings.Configure(x =>
                {
                    x.TableId = "myTable";
                    x.Minify = true;
                    x.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                });
                var html = exporter.GetHtmlString();
                using (var ms = new MemoryStream())
                {
                    exporter.RenderHtml(ms);
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                    var expectedHtml = "<table class=\"epplus-table ts-dark1 ts-dark1-header ts-dark1-row-stripes\" id=\"myTable\"><thead><tr><th data-datatype=\"string\" class=\"epp-al\">Name</th><th data-datatype=\"number\" class=\"epp-al\">Age</th></tr></thead><tbody><tr><td>John Doe</td><td data-value=\"23\" class=\"epp-ar\">23</td></tr></tbody></table>";
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

                var exporter = table.CreateHtmlExporter();
                exporter.Settings.Configure(x =>
                {
                    x.TableId = "myTable";
                    x.Minify = true;
                });

                using (var ms = new MemoryStream())
                {
                    exporter.RenderHtml(ms);
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                    var expectedHtml = "<table class=\"epplus-table ts-dark1 ts-dark1-header ts-dark1-row-stripes\" id=\"myTable\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\" role=\"columnheader\" scope=\"col\">Name</th><th data-datatype=\"number\" class=\"epp-al\" role=\"columnheader\" scope=\"col\">Age</th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\">John Doe</td><td data-value=\"23\" role=\"cell\" class=\"epp-ar\">23</td></tr></tbody></table>";
                    Assert.AreEqual(expectedHtml, result);
                }
            }
        }

        [TestMethod]
        public void ExportAllTableStyles()
        {
            string path = _htmlOutput + "TableStyles";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtml.xlsx", true))
            {
                foreach (TableStyles e in Enum.GetValues(typeof(TableStyles)))
                {
                    if (!(e == TableStyles.Custom || e == TableStyles.None))
                    {
                        var ws = p.Workbook.Worksheets.Add(e.ToString());
                        LoadTestdata(ws);
                        var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                        tbl.TableStyle = e;

                        var exporter = tbl.CreateHtmlExporter();
                        var html = exporter.GetSinglePage();

                        File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                    }
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public async Task ExportAllTableStylesAsync()
        {
            string path = _htmlOutput + "TableStylesAsync";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtml.xlsx", true))
            {
                foreach (TableStyles e in Enum.GetValues(typeof(TableStyles)))
                {
                    if (!(e == TableStyles.Custom || e == TableStyles.None))
                    {
                        var ws = p.Workbook.Worksheets.Add(e.ToString());
                        LoadTestdata(ws);
                        var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{e}");
                        tbl.TableStyle = e;

                        var exporter = tbl.CreateHtmlExporter();
                        var html = await exporter.GetSinglePageAsync();

                        File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                    }
                }
                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void ExportAllFirstLastTableStyles()
        {
            string path = _htmlOutput + "TableStylesFirstLast";
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

                        var exporter = tbl.CreateHtmlExporter();
                        var html = exporter.GetSinglePage();

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
                    ts.FirstRowStripe.Style.Fill.PatternColor.Tint = 0.10;
                    var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{fs}");
                    tbl.StyleName = ts.Name;

                    var exporter = tbl.CreateHtmlExporter();
                    var html = exporter.GetSinglePage();
                    File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public async Task ExportAllCustomTableStylesAsync()
        {
            string path = _htmlOutput + "TableStylesCustomFillsAsync";
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
                    ts.FirstRowStripe.Style.Fill.PatternColor.Tint = 0.10;
                    var tbl = ws.Tables.Add(ws.Cells["A1:D101"], $"tbl{fs}");
                    tbl.StyleName = ts.Name;

                    var exporter = tbl.CreateHtmlExporter();
                    var html = await exporter.GetSinglePageAsync();
                    File.WriteAllText($"{path}\\table-{tbl.StyleName}.html", html);
                }
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ExportAllGradientTableStyles()
        {
            string path = _htmlOutput + "TableStylesGradientFills";
            CreatePathIfNotExists(path);
            using (var p = OpenPackage("TableStylesToHtmlGradientFill.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add($"PatterFill-Gradient");
                LoadTestdata(ws);
                var ts = p.Workbook.Styles.CreateTableStyle($"CustomPattern-Gradient1", TableStyles.Medium9);
                ts.FirstRowStripe.Style.Fill.Style = eDxfFillStyle.GradientFill;
                ts.FirstRowStripe.Style.Fill.Gradient.GradientType = eDxfGradientFillType.Path;
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

                var exporter = tbl.CreateHtmlExporter();
                var html = exporter.GetSinglePage();
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
                LoadTestdata(ws, 100, 1, 1, true);

                var tbl = ws.Tables.Add(ws.Cells["A1:E101"], $"tblGradient");
                tbl.TableStyle = TableStyles.Dark3;
                ws.Cells["A1"].Style.Font.Italic = true;
                ws.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells["C5"].Style.Font.Size = 18;
                tbl.Columns[0].TotalsRowLabel = "Total";
                var exporter = tbl.CreateHtmlExporter();
                var html = exporter.GetSinglePage();
                File.WriteAllText($"{path}\\table-{tbl.StyleName}-CellStyle.html", html);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void ShouldExportWithOtherCultureInfo()
        {
            string path = _htmlOutput + "culture";
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

                var exporter = tbl.CreateHtmlExporter();
                exporter.Settings.Culture = new System.Globalization.CultureInfo("en-US");
                var html = exporter.GetSinglePage();

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
                var exporter = tbl.CreateHtmlExporter();
                exporter.Settings.Configure(x =>
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

                var s = exporter.Settings;
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

                exporter.Settings.ResetToDefault();

                s = exporter.Settings;
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
        [TestMethod]
        public void ShouldExportRichTextAsInlineHtml()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"RichText");

                var rt = ws.Cells["A1"].RichText;
                var rt1 = rt.Add("Header");
                rt1.Color = Color.Red;
                var rt2 = rt.Add(" 1");
                rt2.Color = Color.Blue;

                rt = ws.Cells["B1"].RichText;
                rt1 = rt.Add("Header");
                rt1.Italic = true;
                rt1.Bold = true;
                rt2 = rt.Add(" 2");
                rt2.Strike = true;

                rt = ws.Cells["C1"].RichText;
                rt1 = rt.Add("Header");
                rt1.FontName = "Arial";
                rt1.Size = 12;
                rt2 = rt.Add(" 3");
                rt2.UnderLine = true;

                rt = ws.Cells["A2"].RichText;
                rt1 = rt.Add("Text");
                rt1.Color = Color.Green;
                rt2 = rt.Add(" 1");
                rt2.Color = Color.Yellow;

                rt = ws.Cells["B2"].RichText;
                rt1 = rt.Add("Text");
                rt1.Italic = true;
                rt1.Bold = true;
                rt2 = rt.Add(" 2");
                rt2.Strike = true;

                rt = ws.Cells["C2"].RichText;
                rt1 = rt.Add("Text");
                rt1.FontName = "Times New Roman";
                rt1.Size = 8;
                rt2 = rt.Add(" 3");
                rt2.UnderLine = true;


                var tbl = ws.Tables.Add(ws.Cells["A1:C2"], $"tblRichtext");
                tbl.TableStyle = TableStyles.Dark5;

                var exporter = tbl.CreateHtmlExporter();
                var html = exporter.GetHtmlString();
                var htmlCss = exporter.GetSinglePage();
            }
        }
        [TestMethod]
        public async Task WriteImages_TableAsync()
        {
            using (var p = OpenTemplatePackage("20-CreateAFileSystemReport-Table.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[0];
                var exporter = sheet.Tables[0].CreateHtmlExporter();
                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Pictures.Include = ePictureInclude.Include;
                exporter.Settings.Minify = false;
                var html = exporter.GetSinglePage();
                var htmlAsync = await exporter.GetSinglePageAsync();
                File.WriteAllText($"{_htmlOutput}\\" + sheet.Name + "-table.html", html);
                File.WriteAllText($"{_htmlOutput}\\" + sheet.Name + "-table-async.html", htmlAsync);
                Assert.AreEqual(html, htmlAsync);
            }
        }
        [TestMethod]
        public async Task WriteTableFromRange()
        {
            using (var p = OpenTemplatePackage("20-CreateAFileSystemReport.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[1];
                var exporterRange = sheet.Tables[0].Range.CreateHtmlExporter();
                exporterRange.Settings.SetColumnWidth = true;
                exporterRange.Settings.SetRowHeight = true;
                exporterRange.Settings.Minify = false;
                exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.ClassNamesOnly;
                var html = exporterRange.GetHtmlString();
                var htmlAsync = await exporterRange.GetHtmlStringAsync();

                var css = exporterRange.GetCssString();
                var cssAsync = await exporterRange.GetCssStringAsync();

                var outputHtml = string.Format("<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>", html, css);

                File.WriteAllText($"{_htmlOutput}TableRangeCombined.html", outputHtml);

                Assert.AreEqual(html, htmlAsync);
                Assert.AreEqual(css, cssAsync);
            }
        }
        [TestMethod]
        public async Task WriteMultipleRangeWithTableAndRange()
        {
            using (var p = OpenTemplatePackage("20-CreateAFileSystemReport.xlsx"))
            {
                var sheet1 = p.Workbook.Worksheets[0];
                var sheet2 = p.Workbook.Worksheets[1];

                var exporterRange = p.Workbook.CreateHtmlExporter(
                    sheet2.Tables[0].Range,
                    sheet1.Cells["A1:E30"],
                    sheet2.Tables[2].Range,
                    sheet2.Tables[1].Range);

                exporterRange.Settings.SetColumnWidth = true;
                exporterRange.Settings.SetRowHeight = true;
                exporterRange.Settings.Minify = false;
                exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.Include;
                exporterRange.Settings.Pictures.Include = ePictureInclude.Include;

                var html1 = exporterRange.GetHtmlString(0);
                var html2 = exporterRange.GetHtmlString(1);
                var html3 = exporterRange.GetHtmlString(2);
                var html4 = exporterRange.GetHtmlString(3);

                var css = exporterRange.GetCssString();
                var cssAsync = await exporterRange.GetCssStringAsync();

                var outputHtml = string.Format("<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{4}</style></head>\r\n<body>\r\n{0}<hr>{1}<hr>{2}<hr>{3}<hr></body>\r\n</html>", html1, html2, html3, html4, css);

                File.WriteAllText("${_htmlOutput}RangeAndThreeTables.html", outputHtml);

                Assert.AreEqual(css, cssAsync);
            }
        }
        [TestMethod]
        public async Task WriteAdvancedWs()
        {
            using (var p = OpenTemplatePackage("s610.xlsx"))
            {
                var sheet1 = p.Workbook.Worksheets[0];
                var exporterRange = p.Workbook.CreateHtmlExporter(sheet1.Cells["A1:BL7868"]);
                exporterRange.Settings.SetColumnWidth = true;
                exporterRange.Settings.SetRowHeight = true;
                exporterRange.Settings.Minify = false;
                exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.Include;
                exporterRange.Settings.Pictures.Include = ePictureInclude.Include;
                var htmlAsync = await exporterRange.GetSinglePageAsync();

                File.WriteAllText($"{_htmlOutput}RangeAndThreeTables.html", htmlAsync);
            }
        }
        [TestMethod]
        public async Task Export_CondtionalFormattingHtmlExport_Worksheet1()
        {
            using (var p = OpenTemplatePackage("CondtionalFormattingHtmlExport.xlsx"))
            {
                var sheet1 = p.Workbook.Worksheets[0];
                var exporterRange = p.Workbook.CreateHtmlExporter(sheet1.Cells["A2:F15"]);
                //exporterRange.Settings.SetColumnWidth = true;
                //exporterRange.Settings.SetRowHeight = true;
                //exporterRange.Settings.Minify = false;
                //exporterRange.Settings.TableStyle = eHtmlRangeTableInclude.Include;
                //exporterRange.Settings.Pictures.Include = ePictureInclude.Include;
                var htmlAsync = await exporterRange.GetSinglePageAsync();

                File.WriteAllText($"{_htmlOutput}CondtionalFormattingHtmlExport.html", htmlAsync);
            }
        }
        [TestMethod]
        public void ExportTables()
        {
            using (var p = OpenTemplatePackage("htmlWhiteBorder.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];

                var exporter = ws.Tables[0].CreateHtmlExporter();
                var css = exporter.GetCssString();
                var html = exporter.GetHtmlString();
                var htmlTemplate = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{0}</style></head>\r\n<body>\r\n{1}<hr></body>\r\n</html>";
                var outputHtml = string.Format(htmlTemplate, css, html);
                File.WriteAllText($"{_htmlOutput}TableBorders.html", outputHtml);
    //            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-04-MultipleTables.html", true).FullName,
    //string.Format(htmlTemplate, css, tbl1Html, tbl2Html, tbl3Html));
            }
        }
    }
}
