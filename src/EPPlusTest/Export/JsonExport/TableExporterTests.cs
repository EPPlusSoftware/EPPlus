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
                var tbl = ws.Tables.Add(ws.Cells["A1:G2"], $"tblGradient");

                var s = tbl.ToJson();
                Assert.AreEqual("{\"table\":{\"name\":\"tblGradient\",\"showHeader\":\"1\",\"showTotal\":\"0\",\"column\":[{\"Name\":\"Column1\",\"datatype\":\"string\"},{\"Name\":\"Column2\",\"datatype\":\"string\"},{\"Name\":\"Column3\",\"datatype\":\"string\"},{\"Name\":\"Column4\",\"datatype\":\"string\"},{\"Name\":\"Column5\",\"datatype\":\"string\"},{\"Name\":\"Column6\",\"datatype\":\"string\"},{\"Name\":\"Column7\",\"datatype\":\"string\"}],\"rows\":[{\"cells\":[{\"v\":\"\\\"\",\"t\":\"\\\"\"},{\"v\":\"\\r\\n\",\"t\":\"\\r\\n\"},{\"v\":\"\\f\",\"t\":\"\\f\"},{\"v\":\"\\t\",\"t\":\"\\t\"},{\"v\":\"\\b\",\"t\":\"\\b\"},{\"v\":\"\\t\",\"t\":\"\\t\"},{\"v\":\"\\u0000\",\"t\":\"\\u0000\"}]}]}}"
                    , s);
            }
        }
    }
}
