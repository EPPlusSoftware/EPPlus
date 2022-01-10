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
                    range.HtmlExporter.Settings.FirstRowIsHeader = true;
                    range.HtmlExporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes=false;
                    range.HtmlExporter.RenderHtml(ms);                    
                    var sr = new StreamReader(ms);
                    ms.Position = 0;
                    var result = sr.ReadToEnd();
                    Assert.AreEqual(
                        "<table><thead><tr><th data-datatype=\"string\">Name</th><th data-datatype=\"number\">Age</th></tr></thead><tbody><tr><td>John Doe</td><td data-value=\"23\">23</td></tr><tr><td></td><td></td></tr></tbody></table>",
                        result);
                }
            }
        }
        [TestMethod]
        public void ShouldExportHtmlWithHeadersWithStyles()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Age";
                sheet.Cells["A2"].Value = "John Doe";
                sheet.Cells["B2"].Value = 23;
                var range = sheet.Cells["A1:B2"];
                sheet.Cells["A1:B1"].Style.Font.Bold = true;
                sheet.Cells["A1:B1"].Style.Font.Color.SetColor(Color.Blue);
                sheet.Cells["A1:B1"].Style.Border.Bottom.Style=ExcelBorderStyle.Thin;
                sheet.Cells["A1:B1"].Style.Border.Bottom.Color.SetColor(Color.Red);
                sheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.DarkTrellis;
                sheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                sheet.Cells["A1:B1"].Style.Fill.PatternColor.SetColor(Color.LightCyan);
                sheet.Cells["A2:B2"].Style.Font.Italic=true;
                sheet.Cells["B1:B2"].Style.Font.Name = "Consolas";

                range.HtmlExporter.Settings.FirstRowIsHeader = true;
                range.HtmlExporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes = false;
                var result = range.HtmlExporter.GetSinglePage();
                Assert.AreEqual(
                    "",
                    result);
            }
        }

    }
}
