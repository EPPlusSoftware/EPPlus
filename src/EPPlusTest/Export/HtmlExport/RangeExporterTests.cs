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
    public class RangeExporterTests : TestBase
    {
        [TestMethod]
        public void ShouldExportHtmlWithHeadersNoAccessibilityAttributes()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = 23;
                var range = sheet.Cells["A1:B2"];
                using(var ms = new MemoryStream())
                {
                    var exporter = range.CreateHtmlExporter();
                    exporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes=false;
                    exporter.Settings.Culture = new CultureInfo("us-en");
                    exporter.RenderHtml(ms);                    
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                    Assert.AreEqual(
                        "<table class=\"epplus-table\"><thead><tr><th data-datatype=\"string\" class=\"epp-al\">Name</th><th data-datatype=\"number\" class=\"epp-al\">Age</th></tr></thead><tbody><tr><td>John Doe</td><td data-value=\"23\" class=\"epp-ar\">23</td></tr></tbody></table>",
                        result);
                }
            }
        }
        [TestMethod]
        public void ShouldSetWidthAndDefaultRowAndWidthClasses()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = 23;
                sheet.Cells["A1:A2"].AutoFitColumns();
                var range = sheet.Cells["A1:C3"];

                var exporter = range.CreateHtmlExporter();
                exporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Culture = new CultureInfo("us-en");
                var result = exporter.GetSinglePage();
                Assert.AreEqual(
                    "<!DOCTYPE html><html><head><style type=\"text/css\">table.epplus-table{font-family:Calibri;font-size:11pt;border-spacing:0;border-collapse:collapse;word-wrap:break-word;white-space:nowrap;}.epp-hidden {display:none;}.epp-al {text-align:left;}.epp-ar {text-align:right;}.epp-dcw {width:64px;}.epp-drh {height:20px;}</style></head><body><table class=\"epplus-table\"><colgroup><col class=\"epp-dcw\" span=\"1\"/><col class=\"epp-dcw\" span=\"1\"/><col class=\"epp-dcw\" span=\"1\"/></colgroup><thead><tr class=\"epp-drh\"><th data-datatype=\"string\" class=\"epp-al\">Name</th><th data-datatype=\"number\" class=\"epp-al\">Age</th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody><tr class=\"epp-drh\"><td>John Doe</td><td data-value=\"23\" class=\"epp-ar\">23</td><td></td></tr><tr class=\"epp-drh\"><td></td><td></td><td></td></tr></tbody></table></body></html>",
                    result);
            }
        }

        [TestMethod]
        public async Task ShouldExportHtmlWithHeadersWithStyles()
        {
            using (var package = OpenPackage("HtmlPatternStylesCells.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("PatternStyle");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = 23;
                var range = sheet.Cells["A1:B2"];
                sheet.Cells["A1:B1"].Style.Font.Bold = true;
                sheet.Cells["A1:B1"].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells["A1:B1"].Style.Border.Bottom.Style=ExcelBorderStyle.Thin;
                sheet.Cells["A1:B1"].Style.Border.Bottom.Color.SetColor(Color.Red);
                sheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.LightGray;
                sheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                sheet.Cells["A1:B1"].Style.Fill.PatternColor.SetColor(Color.LightCyan);
                sheet.Cells["A2:B2"].Style.Font.Italic=true;
                sheet.Cells["B1:B2"].Style.Font.Name = "Consolas";

                var exporter = range.CreateHtmlExporter();
                exporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                exporter.Settings.Culture = new CultureInfo("sv-SE");
                var result = exporter.GetSinglePage();

                Assert.AreEqual(
                    "<!DOCTYPE html><html><head><style type=\"text/css\">table.epplus-table{font-family:Calibri;font-size:11pt;border-spacing:0;border-collapse:collapse;word-wrap:break-word;white-space:nowrap;}.epp-hidden {display:none;}.epp-al {text-align:left;}.epp-ar {text-align:right;}.epp-dcw {width:64px;}.epp-drh {height:20px;}.epp-s1{background-repeat:repeat;background:url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPSc0JyBoZWlnaHQ9JzInPjxyZWN0IHdpZHRoPSc0JyBoZWlnaHQ9JzInIGZpbGw9JyNlMGZmZmYnLz48cmVjdCB4PScyJyB5PScwJyB3aWR0aD0nMScgaGVpZ2h0PScxJyBmaWxsPScjZjA4MDgwJy8+PHJlY3QgeD0nMCcgeT0nMScgd2lkdGg9JzEnIGhlaWdodD0nMScgZmlsbD0nI2YwODA4MCcvPjwvc3ZnPg==);color:#0000ff;font-weight:bolder;border-bottom:thin solid #ff0000;white-space: nowrap;}.epp-s2{background-repeat:repeat;background:url(data:image/svg+xml;base64,PHN2ZyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPSc0JyBoZWlnaHQ9JzInPjxyZWN0IHdpZHRoPSc0JyBoZWlnaHQ9JzInIGZpbGw9JyNlMGZmZmYnLz48cmVjdCB4PScyJyB5PScwJyB3aWR0aD0nMScgaGVpZ2h0PScxJyBmaWxsPScjZjA4MDgwJy8+PHJlY3QgeD0nMCcgeT0nMScgd2lkdGg9JzEnIGhlaWdodD0nMScgZmlsbD0nI2YwODA4MCcvPjwvc3ZnPg==);font-family:Consolas;color:#0000ff;font-weight:bolder;border-bottom:thin solid #ff0000;white-space: nowrap;}.epp-s3{font-style:italic;white-space: nowrap;}.epp-s4{font-family:Consolas;font-style:italic;white-space: nowrap;}</style></head><body><table class=\"epplus-table\"><thead><tr><th data-datatype=\"string\" class=\"epp-al epp-s1\">Name</th><th data-datatype=\"number\" class=\"epp-al epp-s2\">Age</th></tr></thead><tbody><tr><td class=\"epp-s3\">John Doe</td><td data-value=\"23\" class=\"epp-ar epp-s4\">23</td></tr></tbody></table></body></html>",
                    result);

                var resultAsync = await exporter.GetSinglePageAsync();
                Assert.AreEqual(result, resultAsync);
                SaveAndCleanup(package);
            }
        }
        [TestMethod]
        public async Task ShouldExportHtmlWithMergedCells()
        {
            using (var package = OpenPackage("HtmlMergeCells.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("Horizontal");
                sheet.Cells["A1"].Value = "Merge Horizontal";
                sheet.Cells["A1:C1"].Merge = true;
                sheet.Cells["C2:C4"].Merge = true;
                sheet.Cells["C2"].Value = "Merge Vertical";
                sheet.Cells["C2"].Style.TextRotation = 255;

                sheet.Cells["A2"].Value = "Name";
                sheet.Cells["B2"].Value = "Age";
                sheet.Cells["A3"].Value = "John Doe";
                sheet.Cells["B3"].Value = 23;
                sheet.Cells["A3"].Value = "Jane Doe";
                sheet.Cells["B3"].Value = 25;
                sheet.Cells["A4"].Value = "James Doe";
                sheet.Cells["B4"].Value = 2;

                sheet.Cells["A1:B1"].Style.Font.Bold = true;
                sheet.Cells["A1:B1"].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells["A1:B1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A1:B1"].Style.Border.Bottom.Color.SetColor(Color.Red);
                sheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.DarkTrellis;
                sheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                sheet.Cells["A1:B1"].Style.Fill.PatternColor.SetColor(Color.LightCyan);
                sheet.Cells["A2:B2"].Style.Font.Italic = true;
                sheet.Cells["B1:B2"].Style.Font.Name = "Consolas";

                var range = sheet.Cells["A1:C4"];
                var exporter = range.CreateHtmlExporter();
                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                var result = exporter.GetSinglePage();
                var resultAsync = await exporter.GetSinglePageAsync();
                SaveAndCleanup(package);
                Assert.AreEqual(result, resultAsync);

            }
        }
        [TestMethod]
        public void WriteHtmlFiles()
        {
            using (var package = OpenTemplatePackage("issue485.xlsx"))
            {
                SaveRangeFile(package, "Avances", "B3:T112");
                SaveRangeFile(package, "Avances TD", "B2:L62");
                SaveRangeFile(package, "Excel docClikalia", "A1:Q345");
            }
            using (var package = OpenTemplatePackage("Calculate Worksheet.xlsx"))
            {
                SaveRangeFile(package, "All Questions", "D1:BF1049", 2);
            }
        }

        [TestMethod]
        public void WriteAllsvenskan()
        {
            using (var p = OpenTemplatePackage("Allsvenskan2001.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[0];
                var exporter = sheet.Cells["B5:N19"].CreateHtmlExporter();
                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Pictures.Include = ePictureInclude.Include;
                var html =exporter.GetSinglePage();
                File.WriteAllText("c:\\temp\\" + sheet.Name + ".html", html);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public async Task WriteImagesAsync()
        {
            using (var p = OpenTemplatePackage("20-CreateAFileSystemReport.xlsx"))
            {
                var sheet = p.Workbook.Worksheets[0];                
                var exporter = sheet.Cells["A1:E30"].CreateHtmlExporter();
                
                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Pictures.Include = ePictureInclude.Include;
                exporter.Settings.Minify = false;
                exporter.Settings.Encoding = Encoding.UTF8; 
                var html = exporter.GetSinglePage();
                var htmlAsync = await exporter.GetSinglePageAsync(); 
                File.WriteAllText("c:\\temp\\" + sheet.Name + ".html", html);
                File.WriteAllText("c:\\temp\\" + sheet.Name + "-async.html", htmlAsync);
                Assert.AreEqual(html, htmlAsync);
            }
        }
        [TestMethod]
        public void ExportMultipleRanges()
        {
            using (var p = OpenTemplatePackage("20-CreateAFileSystemReport.xlsx"))
            {
                var sheet1 = p.Workbook.Worksheets[0];
                var sheet2 = p.Workbook.Worksheets[1];
                
                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B13"],
                    sheet2.Cells["A16:B26"],
                    sheet2.Cells["A29:B42"]);

                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;
                exporter.Settings.Minify = false;
                exporter.Settings.Encoding = Encoding.UTF8;
                var css = exporter.GetCssString();
                var html1 = exporter.GetHtmlString(0);
                var html2 = exporter.GetHtmlString(1);
                var html3 = exporter.GetHtmlString(2);
                var htmlTemplate = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>";
                var page1 = string.Format(htmlTemplate, html1, css);
                var page2 = string.Format(htmlTemplate, html2, css);
                var page3 = string.Format(htmlTemplate, html3, css);
                File.WriteAllText("c:\\temp\\PageSharedCss1.html", page1);
                File.WriteAllText("c:\\temp\\PageSharedCss2.html", page2);
                File.WriteAllText("c:\\temp\\PageSharedCss3.html", page3);
            }
        }

        [TestMethod]
        public void ExportMultipleRangesOverrides_TableId()
        {
            using (var p = new ExcelPackage())
            {
                var sheet1 = p.Workbook.Worksheets.Add("test");
                var sheet2 = p.Workbook.Worksheets.Add("test2");

                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B2"],
                    sheet2.Cells["A16:B18"],
                    sheet2.Cells["A29:B31"]);

                // default table id
                exporter.Settings.TableId = "id";

                // With instance of settings
                var s1 = new ExcelHtmlOverrideExportSettings();
                s1.TableId = "abc";
                var s2 = new ExcelHtmlOverrideExportSettings();
                s2.TableId = "def";
                var html1 = exporter.GetHtmlString(0, s1);
                var html2 = exporter.GetHtmlString(1, s2);
                var html3 = exporter.GetHtmlString(2);

                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"id2\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

                // with lambda
                html1 = exporter.GetHtmlString(0, x => x.TableId = "abc");
                html2 = exporter.GetHtmlString(1, x => x.TableId = "def");
                html3 = exporter.GetHtmlString(2);

                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"id2\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

            }
        }

        [TestMethod]
        public void ExportMultipleRangesOverrides_AdditionalTableClasses()
        {
            using (var p = new ExcelPackage())
            {
                var sheet1 = p.Workbook.Worksheets.Add("test");
                var sheet2 = p.Workbook.Worksheets.Add("test2");

                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B2"],
                    sheet2.Cells["A16:B18"],
                    sheet2.Cells["A29:B31"]);

                // With instance of settings
                var s1 = new ExcelHtmlOverrideExportSettings();
                s1.AdditionalTableClassNames.Add("abc");
                var s2 = new ExcelHtmlOverrideExportSettings();
                s2.AdditionalTableClassNames.Add("def");
                var html1 = exporter.GetHtmlString(0, s1);
                var html2 = exporter.GetHtmlString(1, s2);
                var html3 = exporter.GetHtmlString(2);

                Assert.AreEqual(
                    "<table class=\"epplus-table abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

                // with lambda
                html1 = exporter.GetHtmlString(0, x => x.AdditionalTableClassNames.Add("abc"));
                html2 = exporter.GetHtmlString(1, x => x.AdditionalTableClassNames.Add("def"));
                html3 = exporter.GetHtmlString(2);

                Assert.AreEqual(
                    "<table class=\"epplus-table abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

            }
        }

        [TestMethod]
        public void ExportMultipleRangesOverrides_Accessibility()
        {
            using (var p = new ExcelPackage())
            {
                var sheet1 = p.Workbook.Worksheets.Add("test");
                var sheet2 = p.Workbook.Worksheets.Add("test2");

                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B2"],
                    sheet2.Cells["A16:B18"],
                    sheet2.Cells["A29:B31"]);

                // With instance of settings
                var s1 = new ExcelHtmlOverrideExportSettings();
                s1.Accessibility.TableSettings.AriaLabel = "al1";
                var s2 = new ExcelHtmlOverrideExportSettings();
                s2.Accessibility.TableSettings.AriaLabel = "al2";
                var html1 = exporter.GetHtmlString(0, s1);
                var html2 = exporter.GetHtmlString(1, s2);
                var html3 = exporter.GetHtmlString(2);

                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\" aria-label=\"al1\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\" aria-label=\"al2\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

            }
        }

        [TestMethod]
        public void ExportMultipleRangesOverridesAsync_TableId()
        {
            using (var p = new ExcelPackage())
            {
                var sheet1 = p.Workbook.Worksheets.Add("test");
                var sheet2 = p.Workbook.Worksheets.Add("test2");

                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B2"],
                    sheet2.Cells["A16:B18"],
                    sheet2.Cells["A29:B31"]);

                // default table id
                exporter.Settings.TableId = "id";

                var s1 = new ExcelHtmlOverrideExportSettings();
                s1.TableId = "abc";
                var s2 = new ExcelHtmlOverrideExportSettings();
                s2.TableId = "def";
                var html1 = exporter.GetHtmlStringAsync(0, s1).Result;
                var html2 = exporter.GetHtmlStringAsync(1, s2).Result;
                var html3 = exporter.GetHtmlStringAsync(2).Result;

                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"id2\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

                // with lambda
                html1 = exporter.GetHtmlStringAsync(0, x => x.TableId = "abc").Result;
                html2 = exporter.GetHtmlStringAsync(1, x => x.TableId = "def").Result;
                html3 = exporter.GetHtmlStringAsync(2).Result;

                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" id=\"id2\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);
            }
        }

        [TestMethod]
        public void ExportMultipleRangesOverridesAsync_AdditionalTableClasses()
        {
            using (var p = new ExcelPackage())
            {
                var sheet1 = p.Workbook.Worksheets.Add("test");
                var sheet2 = p.Workbook.Worksheets.Add("test2");

                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B2"],
                    sheet2.Cells["A16:B18"],
                    sheet2.Cells["A29:B31"]);

                // With instance of settings
                var s1 = new ExcelHtmlOverrideExportSettings();
                s1.AdditionalTableClassNames.Add("abc");
                var s2 = new ExcelHtmlOverrideExportSettings();
                s2.AdditionalTableClassNames.Add("def");
                var html1 = exporter.GetHtmlStringAsync(0, s1).Result;
                var html2 = exporter.GetHtmlStringAsync(1, s2).Result;
                var html3 = exporter.GetHtmlStringAsync(2).Result;

                Assert.AreEqual(
                    "<table class=\"epplus-table abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

                // with lambda
                html1 = exporter.GetHtmlStringAsync(0, x => x.AdditionalTableClassNames.Add("abc")).Result;
                html2 = exporter.GetHtmlStringAsync(1, x => x.AdditionalTableClassNames.Add("def")).Result;
                html3 = exporter.GetHtmlStringAsync(2).Result;

                Assert.AreEqual(
                    "<table class=\"epplus-table abc\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table def\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

            }
        }

        [TestMethod]
        public void ExportMultipleRangesOverridesAsync_Accessibility()
        {
            using (var p = new ExcelPackage())
            {
                var sheet1 = p.Workbook.Worksheets.Add("test");
                var sheet2 = p.Workbook.Worksheets.Add("test2");

                var exporter = p.Workbook.CreateHtmlExporter(
                    sheet2.Cells["A1:B2"],
                    sheet2.Cells["A16:B18"],
                    sheet2.Cells["A29:B31"]);

                // With instance of settings
                var s1 = new ExcelHtmlOverrideExportSettings();
                s1.Accessibility.TableSettings.AriaLabel = "al1";
                var s2 = new ExcelHtmlOverrideExportSettings();
                s2.Accessibility.TableSettings.AriaLabel = "al2";
                var html1 = exporter.GetHtmlStringAsync(0, s1).Result;
                var html2 = exporter.GetHtmlStringAsync(1, s2).Result;
                var html3 = exporter.GetHtmlStringAsync(2).Result;

                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\" aria-label=\"al1\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html1
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\" aria-label=\"al2\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html2
                    );
                Assert.AreEqual(
                    "<table class=\"epplus-table\" role=\"table\"><thead role=\"rowgroup\"><tr role=\"row\"><th data-datatype=\"string\" class=\"epp-al\"></th><th data-datatype=\"string\" class=\"epp-al\"></th></tr></thead><tbody role=\"rowgroup\"><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr><tr role=\"row\" scope=\"row\"><td role=\"cell\"></td><td role=\"cell\"></td></tr></tbody></table>",
                    html3);

            }
        }
        [TestMethod]
        public void ExportRangeIssue()
        {
            using (var package = OpenTemplatePackage("componentsource.xlsx"))
            {
                var sheet1 = package.Workbook.Worksheets[0];
                var sheet2 = package.Workbook.Worksheets[1];
                var sheet3 = package.Workbook.Worksheets[2];

                var exporter = package.Workbook.CreateHtmlExporter(sheet1.Cells["B1:E115"], sheet2.Cells["B1:E147"], sheet3.Cells["B1:E105"]);
                exporter.Settings.HyperlinkTarget = "_blank";
                var html1 = exporter.GetHtmlString(0, x =>
                {
                    x.TableId = "americas-toll-free";
                });
                var html2 = exporter.GetHtmlString(1, x =>
                {
                    x.TableId = "emea-toll-free";
                });
                var html3 = exporter.GetHtmlString(2, x =>
                {
                    x.TableId = "asia-toll-free";
                });
                var css = exporter.GetCssString();
                File.WriteAllText("c:\\temp\\html.html",$"<html><head><style type=\"text/css\">{css}</style></head><body>{html1}</body></html>");
            }
        }

        private static void SaveRangeFile(ExcelPackage package, string ws, string address, int headerRows=1)
        {
            var sheet = package.Workbook.Worksheets[ws];
            var range = sheet.Cells[address];
            var exporter = range.CreateHtmlExporter();
            exporter.Settings.SetColumnWidth = true;
            exporter.Settings.HeaderRows = headerRows;
            File.WriteAllText("c:\\temp\\" + sheet.Name + ".html", exporter.GetSinglePage());
        }
    }
}
    