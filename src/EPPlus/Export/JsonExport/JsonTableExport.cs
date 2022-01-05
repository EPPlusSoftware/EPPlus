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

        private JsonTableExportSettings _settings;
        public JsonTableExport(ExcelTable table, JsonTableExportSettings settings) : base(settings)
        {
            _table = table;
            _settings = settings;
        }
        public string Export()
        {
            var sb = new StringBuilder();
            sb.Append($"{{\"{_settings.RootElementName}\":");
            if (_settings.WriteNameAttribute)
            {
                sb.Append($"{{\"name\":\"{JsonEscape(_table.Name)}\",");
            }
            if (_settings.WriteShowHeaderAttribute)
            {
                sb.Append($"\"showHeader\":\"{(_table.ShowHeader ? "1" : "0")}\",");
            }
            if (_settings.WriteShowTotalsAttribute)
            {
                sb.Append($"\"showTotal\":\"{(_table.ShowTotal ? "1" : "0")}\",");
            }
            if (_settings.WriteColumnsElement)
            {
                WriteColumnData(sb);
            }
            WriteCellData(sb, _table.DataRange);
            sb.Append("}");
            return sb.ToString();
        }

        private void WriteColumnData(StringBuilder sb)
        {
            sb.Append($"\"{_settings.ColumnsElementName}\":[");
            for(int i=0;i<_table.Columns.Count;i++)
            {
                if (i > 0) sb.Append(",");
                sb.Append($"{{\"Name\":\"{_table.Columns[i].Name}\"");
                if(_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
                    sb.Append($",\"datatype\":\"{dt}\"");                    
                }
                sb.Append("}");                
            }
            sb.Append("],");
        }
    }
}
