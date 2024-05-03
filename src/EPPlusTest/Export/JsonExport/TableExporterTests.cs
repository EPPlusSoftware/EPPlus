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

namespace EPPlusTest.Export.JsonExport
{
    [TestClass]
    public class TableExporterTests : TestBase
    {
        [TestMethod]
        public void ValidateJsonExport()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"Sheet1");
                LoadTestdata(ws, 100, 1, 1, true, true);
                ws.Cells["A2"].AddComment("Comment in A2");
                var tbl = ws.Tables.Add(ws.Cells["A1:F100"], $"tblGradient");

                var s = tbl.ToJson();
            }
        }
        [TestMethod]
        public async Task ValidateJsonExportRange()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"Sheet1");
                ws.SetValue("A1", "SEK");
                ws.SetValue("B1", "EUR");
                ws.SetValue("C1", "USD");

                ws.SetValue("A2", 1);
                ws.SetValue("B2", 10.35);
                ws.SetValue("C2", 9.51);

                ws.SetValue("A3", 1);
                ws.SetValue("B3", 10.48);
                ws.SetValue("C3", 9.59);

                var json = ws.Cells["A1:C3"].ToJson(x =>
                {
                    x.AddDataTypesOn = eDataTypeOn.OnColumn;
                    x.Culture = new CultureInfo("sv-SE");
                });
                string jsonAsync;
                using (var ms = new MemoryStream())
                {
                    await ws.Cells["A1:C3"].SaveToJsonAsync(ms, x =>
                    {
                        x.AddDataTypesOn = eDataTypeOn.OnColumn;
                        x.Culture = new CultureInfo("sv-SE");
                    });
                    jsonAsync = Encoding.UTF8.GetString(ms.ToArray());
                }
                Assert.AreEqual(json, jsonAsync);
                Assert.AreEqual("{\"range\":{\"columns\":[{\"name\":\"SEK\",\"dt\":\"number\"},{\"name\":\"EUR\",\"dt\":\"number\"},{\"name\":\"USD\",\"dt\":\"number\"}],\"rows\":[{\"cells\":[{\"v\":\"1\",\"t\":\"1\"},{\"v\":\"10.35\",\"t\":\"10,35\"},{\"v\":\"9.51\",\"t\":\"9,51\"}]},{\"cells\":[{\"v\":\"1\",\"t\":\"1\"},{\"v\":\"10.48\",\"t\":\"10,48\"},{\"v\":\"9.59\",\"t\":\"9,59\"}]}]}}",
                    json);
            }
        }
        [TestMethod]
        public void ValidateJsonEncoding()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"Sheet2");
                ws.SetValue(2, 1, "\"");
                ws.SetValue(2, 2, "\r\n");
                ws.SetValue(2, 3, "\f");
                ws.SetValue(2, 4, "\t");
                ws.SetValue(2, 5, "\b");
                ws.SetValue(2, 6, "\t");
                ws.SetValue(2, 7, "\0");
                ws.SetValue(2, 8, "\u001F");
                var range = ws.Cells["A1:G2"];
                var tbl = ws.Tables.Add(ws.Cells["A1:G2"], $"tblGradient");

                var s = tbl.ToJson();
                Assert.AreEqual("{\"table\":{\"name\":\"tblGradient\",\"showHeader\":\"1\",\"showTotal\":\"0\",\"columns\":[{\"name\":\"Column1\",\"dt\":\"string\"},{\"name\":\"Column2\",\"dt\":\"string\"},{\"name\":\"Column3\",\"dt\":\"string\"},{\"name\":\"Column4\",\"dt\":\"string\"},{\"name\":\"Column5\",\"dt\":\"string\"},{\"name\":\"Column6\",\"dt\":\"string\"},{\"name\":\"Column7\",\"dt\":\"string\"}],\"rows\":[{\"cells\":[{\"v\":\"\\\"\",\"t\":\"\\\"\"},{\"v\":\"\\r\\n\",\"t\":\"\\r\\n\"},{\"v\":\"\\f\",\"t\":\"\\f\"},{\"v\":\"\\t\",\"t\":\"\\t\"},{\"v\":\"\\b\",\"t\":\"\\b\"},{\"v\":\"\\t\",\"t\":\"\\t\"},{\"v\":\"\\u0000\",\"t\":\"\\u0000\"}]}]}}"
                    , s);

                s = range.ToJson(x => x.FirstRowIsHeader = false);
                Assert.AreEqual("{\"range\":{\"rows\":[{\"cells\":[{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"}]},{\"cells\":[{\"v\":\"\\\"\",\"t\":\"\\\"\",\"dt\":\"string\"},{\"v\":\"\\r\\n\",\"t\":\"\\r\\n\",\"dt\":\"string\"},{\"v\":\"\\f\",\"t\":\"\\f\",\"dt\":\"string\"},{\"v\":\"\\t\",\"t\":\"\\t\",\"dt\":\"string\"},{\"v\":\"\\b\",\"t\":\"\\b\",\"dt\":\"string\"},{\"v\":\"\\t\",\"t\":\"\\t\",\"dt\":\"string\"},{\"v\":\"\\u0000\",\"t\":\"\\u0000\",\"dt\":\"string\"}]}]}}"
                    , s);
            }
        }

        [TestMethod]
        public async Task TransposedJson()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Id";
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["C1"].Value = 2;
                sheet.Cells["D1"].Value = 3;
                sheet.Cells["E1"].Value = 4;
                sheet.Cells["F1"].Value = 5;
                sheet.Cells["G1"].Value = 6;
                sheet.Cells["A2"].Value = "Name";
                sheet.Cells["B2"].Value = "Scott";
                sheet.Cells["C2"].Value = "Mats";
                sheet.Cells["D2"].Value = "Jimmy";
                sheet.Cells["E2"].Value = "Cameron";
                sheet.Cells["F2"].Value = "Luther";
                sheet.Cells["G2"].Value = "Josh";

                var json = sheet.Cells["A1:G2"].ToJson(x =>
                {
                    x.AddDataTypesOn = eDataTypeOn.OnColumn;
                    x.Culture = new CultureInfo("sv-SE");
                    x.DataIsTransposed = true;
                });

                string jsonAsync;
                using (var ms = new MemoryStream())
                {
                    await sheet.Cells["A1:G2"].SaveToJsonAsync(ms, x =>
                    {
                        x.AddDataTypesOn = eDataTypeOn.OnColumn;
                        x.Culture = new CultureInfo("sv-SE");
                        x.DataIsTransposed = true;
                    });
                    jsonAsync = Encoding.UTF8.GetString(ms.ToArray());
                }
                Assert.AreEqual(json, jsonAsync);
            }
        }
    }
}
