using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml
{
    internal abstract class JsonExport
    {
        internal protected void WriteCellData(StringBuilder sb, ExcelRangeBase dr)
        {
            ExcelWorksheet ws = dr.Worksheet;
            Uri uri = null;
            int commentIx = 0;
            sb.Append("\"rows\":[");
            for (int r = dr._fromRow; r <= dr._toRow; r++)
            {
                if (r > dr._fromRow) sb.Append(",");
                sb.Append("{\"cells\":[");
                for (int c = dr._fromCol; c <= dr._toCol; c++)
                {
                    if (c > dr._fromCol) sb.Append(",");
                    var cv = ws.GetCoreValueInner(r, c);
                    var t = JsonEscape(ValueToTextHandler.GetFormattedText(cv._value, ws.Workbook, cv._styleId, false));
                    if (cv._value == null)
                    {
                        sb.Append($"{{\"t\":\"{t}\"");
                    }
                    else
                    {
                        var v = JsonEscape(HtmlRawDataProvider.GetRawValue(cv._value));
                        sb.Append($"{{\"v\":\"{v}\",\"t\":\"{t}\"");
                    }

                    if (ws._hyperLinks.Exists(r, c, ref uri))
                    {
                        sb.Append($",\"uri\":\"{JsonEscape(uri?.OriginalString)}\"");
                    }

                    if (ws._commentsStore.Exists(r, c, ref commentIx))
                    {
                        var comment = ws.Comments[commentIx];
                        sb.Append($",\"comment\":\"{comment.Text}\"");
                    }

                    sb.Append("}");
                }
                sb.Append("]}");
            }
            sb.Append("]");
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
