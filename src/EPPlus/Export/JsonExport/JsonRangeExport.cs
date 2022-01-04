using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal class JsonRangeExport : JsonExport
    {
        private ExcelRangeBase _range;

        public JsonRangeExport(ExcelRangeBase range)
        {
            _range = range;
        }
        public string Export()
        {
            var sb = new StringBuilder();
            sb.Append($"{{\"range\":");
            //WriteColumnData(sb);
            WriteCellData(sb, _range);
            sb.Append("}}");
            return sb.ToString();
        }

        //private void WriteColumnData(StringBuilder sb)
        //{
        //    sb.Append("\"columns\":[");
        //    for(int i=0;i<_table.Columns.Count;i++)
        //    {
        //        if (i > 0) sb.Append(",");
        //        var dt =HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
        //        sb.Append($"{{\"Name\":\"{_table.Columns[i].Name}\",\"datatype\":\"{dt}\"}}");
        //    }
        //    sb.Append("],");
        //}
    }
}
