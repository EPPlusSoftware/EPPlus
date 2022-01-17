using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal partial class JsonTableExport : JsonExport
    {
        private ExcelTable _table;

        private JsonTableExportSettings _settings;
        internal JsonTableExport(ExcelTable table, JsonTableExportSettings settings) : base(settings)
        {
            _table = table;
            _settings = settings;
        }
        internal void Export(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            sw.Write($"{{\"{_settings.RootElementName}\":{{");
            if (_settings.WriteNameAttribute)
            {
                sw.Write($"\"name\":\"{JsonEscape(_table.Name)}\",");
            }
            if (_settings.WriteShowHeaderAttribute)
            {
                sw.Write($"\"showHeader\":\"{(_table.ShowHeader ? "1" : "0")}\",");
            }
            if (_settings.WriteShowTotalsAttribute)
            {
                sw.Write($"\"showTotal\":\"{(_table.ShowTotal ? "1" : "0")}\",");
            }
            if (_settings.WriteColumnsElement)
            {
                WriteColumnData(sw);
            }
            WriteCellData(sw, _table.DataRange);
            sw.Write("}");
            sw.Flush();
        }

        private void WriteColumnData(StreamWriter sw)
        {
            sw.Write($"\"{_settings.ColumnsElementName}\":[");
            for(int i=0;i<_table.Columns.Count;i++)
            {
                if (i > 0) sw.Write(",");
                sw.Write($"{{\"Name\":\"{_table.Columns[i].Name}\"");
                if(_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
                    sw.Write($",\"datatype\":\"{dt}\"");                    
                }
                sw.Write("}");                
            }
            sw.Write("],");
        }
    }
}
