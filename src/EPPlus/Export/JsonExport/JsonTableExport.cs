using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal class JsonTableExport : JsonExport
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
            sb.Append($"\"showHeader\":\"{(_table.ShowHeader ? "1" : "0")}\",");
            sb.Append($"\"showTotal\":\"{(_table.ShowTotal ? "1" : "0")}\",");
            WriteColumnData(sb);
            WriteCellData(sb, _table.DataRange);
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
    }
}
