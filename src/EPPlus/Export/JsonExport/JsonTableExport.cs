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
            WriteStart(sw);
            WriteItem(sw, $"\"{_settings.RootElementName}\":");
            WriteStart(sw);
            if (_settings.WriteNameAttribute)
            {
                WriteItem(sw, $"\"name\":\"{JsonEscape(_table.Name)}\",");
            }
            if (_settings.WriteShowHeaderAttribute)
            {
                WriteItem(sw, $"\"showHeader\":\"{(_table.ShowHeader ? "1" : "0")}\",");
            }
            if (_settings.WriteShowTotalsAttribute)
            {
                WriteItem(sw, $"\"showTotal\":\"{(_table.ShowTotal ? "1" : "0")}\",");
            }
            if (_settings.WriteColumnsElement)
            {
                WriteColumnData(sw);
            }
            WriteCellData(sw, _table.DataRange, 0);
            sw.Write("}");
            sw.Flush();
        }

        private void WriteColumnData(StreamWriter sw)
        {
            WriteItem(sw, $"\"{_settings.ColumnsElementName}\":[", true);
            for (int i = 0; i < _table.Columns.Count; i++)
            {
                WriteStart(sw);
                WriteItem(sw, $"\"name\":\"{JsonEscape(_table.Columns[i].Name)}\"", false, _settings.AddDataTypesOn == eDataTypeOn.OnColumn);
                if (_settings.AddDataTypesOn == eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
                    WriteItem(sw, $"\"dt\":\"{dt}\"");
                }
                if(i == _table.Columns.Count-1)
                {
                    WriteEnd(sw, "}");
                }
                else
                {
                    WriteEnd(sw, "},");
                }
            }
            WriteEnd(sw, "],");
        }
    }
}
