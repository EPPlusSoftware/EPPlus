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
    internal partial class JsonRangeExport : JsonExport
    {
        private ExcelRangeBase _range;
        private JsonRangeExportSettings _settings;
        internal JsonRangeExport(ExcelRangeBase range, JsonRangeExportSettings settings) : base(settings)
        {
            _range = range;
            _settings = settings;
        }
        internal void Export(Stream stream)
        {
            var sw = new StreamWriter(stream);
            WriteStart(sw);
            WriteItem(sw, $"\"{_settings.RootElementName}\":");
            WriteStart(sw);
            if (_settings.FirstRowIsHeader || (_settings.AddDataTypesOn==eDataTypeOn.OnColumn && _range.Rows>1))
            {
                WriteColumnData(sw);
            }
            WriteCellData(sw, _range, _settings.FirstRowIsHeader ? 1 : 0);
            sw.Write("}");
            sw.Flush();
        }

        private void WriteColumnData(StreamWriter sw)
        {
            WriteItem(sw, $"\"{_settings.ColumnsElementName}\":[", true);
            for (int i = 0; i < _range.Columns; i++)
            {
                WriteStart(sw);
                if (_settings.FirstRowIsHeader)
                {
                    WriteItem(sw, $"\"name\":\"{JsonEscape(_range.GetCellValue<string>(0,i))}\"", false, _settings.AddDataTypesOn == eDataTypeOn.OnColumn);
                }
                if (_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(1, i));
                    WriteItem(sw, $"\"dt\":\"{dt}\"");
                }
                if (i == _range.Columns - 1)
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
