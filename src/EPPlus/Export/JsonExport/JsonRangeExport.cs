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
        private JsonRangeExportSettings _settings;
        public JsonRangeExport(ExcelRangeBase range, JsonRangeExportSettings settings) : base(settings)
        {
            _range = range;
            _settings = settings;
        }
        public string Export()
        {
            var sb = new StringBuilder();
            sb.Append($"{{\"{_settings.RootElementName}\":");
            if (_settings.FirstRowIsHeader || (_settings.AddDataTypesOn==eDataTypeOn.OnColumn && _range.Rows>1))
            {
                WriteColumnData(sb);
            }
            WriteCellData(sb, _range);
            sb.Append("}}");
            return sb.ToString();
        }

        private void WriteColumnData(StringBuilder sb)
        {
            sb.Append($"\"{_settings.ColumnsElementName}\":[");
            for (int i = 0; i < _range.Columns; i++)
            {
                if (i > 0) sb.Append(",");
                sb.Append("{");
                if (_settings.FirstRowIsHeader)
                {
                    sb.Append($"\"Name\":\"{_range.GetCellValue<string>(0,i)}\"");
                }
                if (_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    if (_settings.FirstRowIsHeader) sb.Append(",");
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(1, i));
                    sb.Append($"\"dataType\":\"{dt}\"");
                }
                sb.Append("}");
            }


            sb.Append("],");
        }
    }
}
