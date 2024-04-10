using OfficeOpenXml.Export.HtmlExport;
using System.IO;

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
            var total1 = _range.Columns;
            if (_settings.DataIsTransposed)
            {
                total1 = _range.Rows;
            }
            WriteItem(sw, $"\"{_settings.ColumnsElementName}\":[", true);
            for (int i = 0; i < total1; i++)
            {
                WriteStart(sw);
                if (_settings.FirstRowIsHeader)
                {
                    var v = _settings.DataIsTransposed ? _range.GetCellValue<string>(i, 0) : _range.GetCellValue<string>(0, i);
                    WriteItem(sw, $"\"name\":\"{JsonEscape(v)}\"", false, _settings.AddDataTypesOn == eDataTypeOn.OnColumn);
                }
                if (_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    var dt = _settings.DataIsTransposed ? HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(i, 1)) : HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(1, i));
                    WriteItem(sw, $"\"dt\":\"{dt}\"");
                }
                if (i == total1 - 1)
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
