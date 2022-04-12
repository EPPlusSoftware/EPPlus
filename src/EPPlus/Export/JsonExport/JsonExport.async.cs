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
        internal protected async Task WriteCellDataAsync(StreamWriter sw, ExcelRangeBase dr, int headerRows)
        {
            bool dtOnCell = _settings.AddDataTypesOn == eDataTypeOn.OnCell;
            ExcelWorksheet ws = dr.Worksheet;
            Uri uri = null;
            int commentIx = 0;
            await WriteItemAsync(sw, $"\"{_settings.RowsElementName}\":[", true);
            var fromRow = dr._fromRow + headerRows;
            for (int r = fromRow; r <= dr._toRow; r++)
            {
                await WriteStartAsync(sw);
                await WriteItemAsync(sw, $"\"{_settings.CellsElementName}\":[", true);
                for (int c = dr._fromCol; c <= dr._toCol; c++)
                {
                    var cv = ws.GetCoreValueInner(r, c);
                    var t = JsonEscape(ValueToTextHandler.GetFormattedText(cv._value, ws.Workbook, cv._styleId, false, _settings.Culture));
                    await WriteStartAsync(sw);
                    var hasHyperlink = _settings.WriteHyperlinks && ws._hyperLinks.Exists(r, c, ref uri);
                    var hasComment = _settings.WriteComments && ws._commentsStore.Exists(r, c, ref commentIx);
                    if (cv._value == null)
                    {
                        await WriteItemAsync(sw, $"\"t\":\"{t}\"");
                    }
                    else
                    {
                        var v = JsonEscape(HtmlRawDataProvider.GetRawValue(cv._value));
                        await WriteItemAsync(sw, $"\"v\":\"{v}\",");
                        await WriteItemAsync(sw, $"\"t\":\"{t}\"", false, dtOnCell || hasHyperlink || hasComment);
                        if (dtOnCell)
                        {
                            var dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cv._value);
                            await WriteItemAsync(sw, $"\"dt\":\"{dt}\"", false, hasHyperlink || hasComment);
                        }
                    }

                    if (hasHyperlink)
                    {
                        await WriteItemAsync(sw, $"\"uri\":\"{JsonEscape(uri?.OriginalString)}\"", false, hasComment);
                    }

                    if (hasComment)
                    {
                        var comment = ws.Comments[commentIx];
                        await WriteItemAsync(sw, $"\"comment\":\"{comment.Text}\"");
                    }

                    if (c == dr._toCol)
                    {
                        await WriteEndAsync(sw, "}");
                    }
                    else
                    {
                        await WriteEndAsync(sw, "},");
                    }
                }
                await WriteEndAsync(sw, "]");
                if (r == dr._toRow)
                {
                    await WriteEndAsync(sw);
                }
                else
                {
                    await WriteEndAsync(sw, "},");
                }
            }
            await WriteEndAsync(sw, "]");
            await WriteEndAsync(sw);
        }
        internal protected async Task WriteItemAsync(StreamWriter sw, string v, bool indent = false, bool addComma = false)
        {
            if (addComma) v += ",";
            if (_minify)
            {
                await sw.WriteAsync(v);
            }
            else
            {
                await sw.WriteLineAsync(_indent + v);
                if (indent)
                {
                    _indent += "  ";
                }
            }
        }

        internal protected async Task WriteStartAsync(StreamWriter sw)
        {
            if (_minify)
            {
                await sw.WriteAsync("{");
            }
            else
            {
                await sw.WriteLineAsync($"{_indent}{{");
                _indent += "  ";
            }
        }
        internal protected async Task WriteEndAsync(StreamWriter sw, string bracket = "}")
        {
            if (_minify)
            {
                await sw.WriteAsync(bracket);
            }
            else
            {
                _indent = _indent.Substring(0, _indent.Length - 2);
                await sw.WriteLineAsync($"{_indent}{bracket}");
            }
        }
    }
}
#endif
