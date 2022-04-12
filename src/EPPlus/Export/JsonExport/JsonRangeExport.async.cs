using OfficeOpenXml.Export.HtmlExport;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml
{
    internal partial class JsonRangeExport : JsonExport
    {
        internal async Task ExportAsync(Stream stream)
        {
            var sw = new StreamWriter(stream);
            await WriteStartAsync(sw);
            await WriteItemAsync(sw, $"\"{_settings.RootElementName}\":");
            await WriteStartAsync(sw);
            if (_settings.FirstRowIsHeader || (_settings.AddDataTypesOn == eDataTypeOn.OnColumn && _range.Rows > 1))
            {
                await WriteColumnDataAsync(sw);
            }
            await WriteCellDataAsync(sw, _range, _settings.FirstRowIsHeader ? 1 : 0);
            await sw.WriteAsync("}");
            await sw.FlushAsync();
        }

        private async Task WriteColumnDataAsync(StreamWriter sw)
        {
            await WriteItemAsync(sw, $"\"{_settings.ColumnsElementName}\":[", true);
            for (int i = 0; i < _range.Columns; i++)
            {
                await WriteStartAsync(sw);
                if (_settings.FirstRowIsHeader)
                {
                    await WriteItemAsync(sw, $"\"name\":\"{_range.GetCellValue<string>(0, i)}\"", false, _settings.AddDataTypesOn == eDataTypeOn.OnColumn);
                }
                if (_settings.AddDataTypesOn == eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(1, i));
                    await WriteItemAsync(sw, $"\"dt\":\"{dt}\"");
                }
                if (i == _range.Columns - 1)
                {
                    await WriteEndAsync(sw, "}");
                }
                else
                {
                    await WriteEndAsync(sw, "},");
                }
            }

            await WriteEndAsync(sw, "],");
        }
    }
}
#endif