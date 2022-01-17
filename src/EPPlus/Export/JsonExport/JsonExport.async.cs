using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
namespace OfficeOpenXml
{
    internal abstract partial class JsonExport
    {
        internal protected async Task WriteCellDataAsync(StreamWriter sw, ExcelRangeBase dr)
        {
            ExcelWorksheet ws = dr.Worksheet;
            Uri uri = null;
            int commentIx = 0;
            await sw.WriteAsync($"\"{_settings.RowsElementName}\":[");
            for (int r = dr._fromRow; r <= dr._toRow; r++)
            {
                if (r > dr._fromRow) sw.Write(",");
                await sw.WriteAsync($"{{\"{_settings.CellsElementName}\":[");
                for (int c = dr._fromCol; c <= dr._toCol; c++)
                {
                    if (c > dr._fromCol) sw.Write(",");
                    var cv = ws.GetCoreValueInner(r, c);
                    var t = JsonEscape(ValueToTextHandler.GetFormattedText(cv._value, ws.Workbook, cv._styleId, false));
                    if (cv._value == null)
                    {
                        await sw.WriteAsync($"{{\"t\":\"{t}\"");
                    }
                    else
                    {
                        var v = JsonEscape(HtmlRawDataProvider.GetRawValue(cv._value));
                        await sw.WriteAsync($"{{\"v\":\"{v}\",\"t\":\"{t}\"");
                        if(_settings.AddDataTypesOn==eDataTypeOn.OnCell)
                        {
                            var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cv._value);
                            await sw.WriteAsync($",\"dataType\":\"{dt}\"");
                        }
                    }

                    if (_settings.WriteHyperlinks && ws._hyperLinks.Exists(r, c, ref uri))
                    {
                        await sw.WriteAsync($",\"uri\":\"{JsonEscape(uri?.OriginalString)}\"");
                    }

                    if (_settings.WriteComments && ws._commentsStore.Exists(r, c, ref commentIx))
                    {
                        var comment = ws.Comments[commentIx];
                        await sw.WriteAsync($",\"comment\":\"{comment.Text}\"");
                    }

                    await sw.WriteAsync("}");
                }
                await sw.WriteAsync("]}");
            }
            await sw.WriteAsync("]}");
        }
    }
}
#endif
