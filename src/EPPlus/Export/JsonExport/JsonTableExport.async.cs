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
            await sw.WriteAsync($"{{\"{_settings.RootElementName}\":{{");
            if (_settings.WriteNameAttribute)
            {
                await sw.WriteAsync($"\"name\":\"{JsonEscape(_table.Name)}\",");
            }
            if (_settings.WriteShowHeaderAttribute)
            {
                await sw.WriteAsync($"\"showHeader\":\"{(_table.ShowHeader ? "1" : "0")}\",");
            }
            if (_settings.WriteShowTotalsAttribute)
            {
                await sw.WriteAsync($"\"showTotal\":\"{(_table.ShowTotal ? "1" : "0")}\",");
            }
            if (_settings.WriteColumnsElement)
            {
                WriteColumnData(sw);
            }
            await WriteCellDataAsync(sw, _table.DataRange);
            await sw.WriteAsync("}");
            await sw.FlushAsync();
        }

        private async Task WriteColumnDataAsync(StreamWriter sw)
        {
            await sw.WriteAsync($"\"{_settings.ColumnsElementName}\":[");
            for(int i=0;i<_table.Columns.Count;i++)
            {
                if (i > 0) sw.Write(",");
                await sw.WriteAsync($"{{\"Name\":\"{_table.Columns[i].Name}\"");
                if(_settings.AddDataTypesOn==eDataTypeOn.OnColumn)
                {
                    var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(_table.DataRange.GetCellValue<object>(0, i));
                    await sw.WriteAsync($",\"datatype\":\"{dt}\"");                    
                }
                await sw.WriteAsync("}");                
            }
            await sw.WriteAsync("],");
        }
    }
}
#endif