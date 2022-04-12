using OfficeOpenXml.Export.HtmlExport;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;
namespace OfficeOpenXml
{
    internal partial class JsonTableExport : JsonExport
    {
        internal async Task ExportAsync(Stream stream)
        {
            StreamWriter sw = new StreamWriter(stream);
            await WriteStartAsync(sw);
            await WriteItemAsync(sw, $"\"{_settings.RootElementName}\":");
            await WriteStartAsync(sw);
            if (_settings.WriteNameAttribute)
            {
                await WriteItemAsync(sw, $"\"name\":\"{JsonEscape(_table.Name)}\",");
            }
            if (_settings.WriteShowHeaderAttribute)
            {
                await WriteItemAsync(sw, $"\"showHeader\":\"{(_table.ShowHeader ? "1" : "0")}\",");
            }
            if (_settings.WriteShowTotalsAttribute)
            {
                await WriteItemAsync(sw, $"\"showTotal\":\"{(_table.ShowTotal ? "1" : "0")}\",");
            }
            if (_settings.WriteColumnsElement)
            {
                await WriteColumnDataAsync(sw);
            }
            await WriteCellDataAsync(sw, _table.DataRange, 0);
            await sw.WriteAsync("}");
            await sw.FlushAsync();
        }

        private async Task WriteColumnDataAsync(StreamWriter sw)
        {
            await WriteItemAsync(sw, $"\"{_settings.ColumnsElementName}\":[", true);
            for (int i = 0; i < _table.Columns.Count; i++)
            {
                await WriteStartAsync(sw);
                await WriteItemAsync(sw, $"\"name\":\"{_table.Columns[i].Name}\"", false, _settings.AddDataTypesOn == eDataTypeOn.OnColumn);
                if (_settings.AddDataTypesOn == eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
                    await WriteItemAsync(sw, $"\"dt\":\"{dt}\"");
                }
                if (i == _table.Columns.Count - 1)
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