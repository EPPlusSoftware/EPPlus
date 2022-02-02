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
            await sw.WriteAsync($"{{\"{_settings.RootElementName}\":{{");
            if (_settings.FirstRowIsHeader || (_settings.AddDataTypesOn==eDataTypeOn.OnColumn && _range.Rows>1))
            {
                await WriteColumnDataAsync(sw);
            }
            await WriteCellDataAsync(sw, _range, _settings.FirstRowIsHeader ? 1 : 0);
            await sw.WriteAsync("}");
            await sw.FlushAsync();
        }

        private async Task WriteColumnDataAsync(StreamWriter sw)
        {
            await sw.WriteAsync($"\"{_settings.ColumnsElementName}\":[");
            for (int i = 0; i < _range.Columns; i++)
            {
                if (i > 0) await sw.WriteAsync(",");
                await sw.WriteAsync("{");
                if (_settings.FirstRowIsHeader)
                {
                    await sw.WriteAsync($"\"Name\":\"{_range.GetCellValue<string>(0,i)}\"");
                }
                if (_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    if (_settings.FirstRowIsHeader) sw.Write(",");
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_range.GetCellValue<object>(1, i));
                    await sw.WriteAsync($"\"dataType\":\"{dt}\"");
                }
                await sw.WriteAsync("}");
            }


            await sw.WriteAsync("],");
        }
    }
}
#endif