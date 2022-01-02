using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    public class JsonTableExport
    {
        private ExcelTable _table;

        public JsonTableExport(ExcelTable table)
        {
            _table = table;
        }
        public string Export()
        {
            var sb = new StringBuilder();
            sb.Append($"{{\"table\":{{\"name\":\"{JsonEscape(_table.Name)}\",");
            WriteColumnData(sb);
            WriteCellData(sb);
            sb.Append("}}");
            return sb.ToString();
        }

        private void WriteColumnData(StringBuilder sb)
        {
            sb.Append("\"columns\":[");
            for(int i=0;i<_table.Columns.Count;i++)
            {
                if (i > 0) sb.Append(",");
                var dt =HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
                sb.Append($"{{\"Name\":\"{_table.Columns[i].Name}\",\"datatype\":\"{dt}\"}}");
            }
            sb.Append("],");
        }

        private void WriteCellData(StringBuilder sb)
        {
            var ws = _table.WorkSheet;
            var dr = _table.DataRange;
            sb.Append("\"rows\":[");
            for (int r=dr._fromRow;r<=dr._toRow;r++)
            {
                if (r > dr._fromRow) sb.Append(",");
                sb.Append("{\"cells\":[");
                for (int c = dr._fromCol; c <= dr._toCol; c++)
                {
                    if (c > dr._fromCol) sb.Append(",");
                    var cv = ws.GetCoreValueInner(r, c);
                    var t = ValueToTextHandler.GetFormattedText(cv._value, _table.WorkSheet.Workbook, cv._styleId, false);
                    if (cv._value == null) 
                    {
                        sb.Append($"{{\"t\":\"{t}\"}}");
                    }
                    else
                    {
                        var v = HtmlRawDataProvider.GetRawValue(cv._value);
                        sb.Append($"{{\"v\":\"{v}\",\"t\":\"{t}\"}}");
                    }
                }
                sb.Append("]}");
            }
            sb.Append("]");
        }
        internal static string JsonEscape(string s)
        {
            return s.Replace("\\", "\\\\").
                     Replace("\"", "\\\"").
                     Replace("\b", "\\b").
                     Replace("\f", "\\f").
                     Replace("\n", "\\n").
                     Replace("\r", "\\r").
                     Replace("\t", "\\t");
        }
    }
}
