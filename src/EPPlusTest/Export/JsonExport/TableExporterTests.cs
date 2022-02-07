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
                });
                string jsonAsync;
                using (var ms = new MemoryStream())
                {
                      await ws.Cells["A1:C3"].SaveToJsonAsync(ms, x =>
                      {
                          x.AddDataTypesOn = eDataTypeOn.OnColumn;
                      });
                    jsonAsync = Encoding.UTF8.GetString(ms.ToArray());
                }
                Assert.AreEqual(json, jsonAsync); 
                Assert.AreEqual("{\"range\":{\"column\":[{\"Name\":\"SEK\",\"dataType\":\"number\"},{\"Name\":\"EUR\",\"dataType\":\"number\"},{\"Name\":\"USD\",\"dataType\":\"number\"}],\"rows\":[{\"cells\":[{\"v\":\"1\",\"t\":\"1\"},{\"v\":\"10.35\",\"t\":\"10,35\"},{\"v\":\"9.51\",\"t\":\"9,51\"}]},{\"cells\":[{\"v\":\"1\",\"t\":\"1\"},{\"v\":\"10.48\",\"t\":\"10,48\"},{\"v\":\"9.59\",\"t\":\"9,59\"}]}]}}", 
                    json);
            }
        }
        [TestMethod]
        public void ValidateJsonEncoding()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add($"Sheet2");
                ws.SetValue(2, 1,"\"");
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
                Assert.AreEqual("{\"table\":{\"name\":\"tblGradient\",\"showHeader\":\"1\",\"showTotal\":\"0\",\"column\":[{\"Name\":\"Column1\",\"datatype\":\"string\"},{\"Name\":\"Column2\",\"datatype\":\"string\"},{\"Name\":\"Column3\",\"datatype\":\"string\"},{\"Name\":\"Column4\",\"datatype\":\"string\"},{\"Name\":\"Column5\",\"datatype\":\"string\"},{\"Name\":\"Column6\",\"datatype\":\"string\"},{\"Name\":\"Column7\",\"datatype\":\"string\"}],\"rows\":[{\"cells\":[{\"v\":\"\\\"\",\"t\":\"\\\"\"},{\"v\":\"\\r\\n\",\"t\":\"\\r\\n\"},{\"v\":\"\\f\",\"t\":\"\\f\"},{\"v\":\"\\t\",\"t\":\"\\t\"},{\"v\":\"\\b\",\"t\":\"\\b\"},{\"v\":\"\\t\",\"t\":\"\\t\"},{\"v\":\"\\u0000\",\"t\":\"\\u0000\"}]}]}}"
                    , s);

                s = range.ToJson(x => x.FirstRowIsHeader = false);
                Assert.AreEqual("{\"range\":{\"rows\":[{\"cells\":[{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"},{\"t\":\"\"}]},{\"cells\":[{\"v\":\"\\\"\",\"t\":\"\\\"\",\"dataType\":\"string\"},{\"v\":\"\\r\\n\",\"t\":\"\\r\\n\",\"dataType\":\"string\"},{\"v\":\"\\f\",\"t\":\"\\f\",\"dataType\":\"string\"},{\"v\":\"\\t\",\"t\":\"\\t\",\"dataType\":\"string\"},{\"v\":\"\\b\",\"t\":\"\\b\",\"dataType\":\"string\"},{\"v\":\"\\t\",\"t\":\"\\t\",\"dataType\":\"string\"},{\"v\":\"\\u0000\",\"t\":\"\\u0000\",\"dataType\":\"string\"}]}]}}}"
                    , s);
            }
        }
    }
}
