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
        public JsonRangeExport(ExcelRangeBase range, JsonRangeExportSettings settings) : base(settings)
        {
            _range = range;
            _settings = settings;
        }
        internal void Export(Stream stream)
        {
            var sw = new StreamWriter(stream);
            sw.Write($"{{\"{_settings.RootElementName}\":{{");
            if (_settings.FirstRowIsHeader || (_settings.AddDataTypesOn==eDataTypeOn.OnColumn && _range.Rows>1))
            {
                WriteColumnData(sw);
            }
            WriteCellData(sw, _range);
            sw.Write("}}");
            sw.Flush();
        }

        private void WriteColumnData(StreamWriter sw)
        {
            sw.Write($"\"{_settings.ColumnsElementName}\":[");
            for (int i = 0; i < _range.Columns; i++)
            {
                if (i > 0) sw.Write(",");
                sw.Write("{");
                if (_settings.FirstRowIsHeader)
                {
                    sw.Write($"\"Name\":\"{_range.GetCellValue<string>(0,i)}\"");
                }
                if (_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    if (_settings.FirstRowIsHeader) sw.Write(",");
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(1, i));
                    sw.Write($"\"dataType\":\"{dt}\"");
                }
                sw.Write("}");
            }


            sw.Write("],");
        }
    }
}
