using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Text;

namespace OfficeOpenXml
{
    internal abstract partial class JsonExport
    {
        private JsonExportSettings _settings;
        public JsonExport(JsonExportSettings settings)
        {
            _settings = settings;
        }
        internal protected void WriteCellData(StreamWriter sw, ExcelRangeBase dr, int headerRows)
        {
            ExcelWorksheet ws = dr.Worksheet;
            Uri uri = null;
            int commentIx = 0;
            sw.Write($"\"{_settings.RowsElementName}\":[");
            var fromRow = dr._fromRow + headerRows;
            for (int r = fromRow; r <= dr._toRow; r++)
            {
                if (r > fromRow) sw.Write(",");
                sw.Write($"{{\"{_settings.CellsElementName}\":[");
                for (int c = dr._fromCol; c <= dr._toCol; c++)
                {
                    if (c > dr._fromCol) sw.Write(",");
                    var cv = ws.GetCoreValueInner(r, c);
                    var t = JsonEscape(ValueToTextHandler.GetFormattedText(cv._value, ws.Workbook, cv._styleId, false));
                    if (cv._value == null)
                    {
                        sw.Write($"{{\"t\":\"{t}\"");
                    }
                    else
                    {
                        var v = JsonEscape(HtmlRawDataProvider.GetRawValue(cv._value));
                        sw.Write($"{{\"v\":\"{v}\",\"t\":\"{t}\"");
                        if(_settings.AddDataTypesOn==eDataTypeOn.OnCell)
                        {
                            var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cv._value);
                            sw.Write($",\"dataType\":\"{dt}\"");
                        }
                    }

                    if (_settings.WriteHyperlinks && ws._hyperLinks.Exists(r, c, ref uri))
                    {
                        sw.Write($",\"uri\":\"{JsonEscape(uri?.OriginalString)}\"");
                    }

                    if (_settings.WriteComments && ws._commentsStore.Exists(r, c, ref commentIx))
                    {
                        var comment = ws.Comments[commentIx];
                        sw.Write($",\"comment\":\"{comment.Text}\"");
                    }

                    sw.Write("}");
                }
                sw.Write("]}");
            }
            sw.Write("]}");
        }
        internal static string JsonEscape(string s)
        {
            if (s == null) return "";
            var sb = new StringBuilder();
            foreach (var c in s)
            {
                switch (c)
                {
                    case '\\':
                        sb.Append("\\\\");
                        break;
                    case '"':
                        sb.Append("\\\"");
                        break;
                    case '\b':
                        sb.Append("\\b");
                        break;
                    case '\f':
                        sb.Append("\\f");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    default:
                        if (c < 0x20)
                        {
                            sb.Append($"\\u{((short)c):X4}");
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            return sb.ToString();
        }
    }
}
