/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    public partial class TableExporter
    {
#if !NET35 && !NET40
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            return await GetHtmlStringAsync(HtmlTableExportOptions.Default);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="options"><see cref="HtmlTableExportOptions">Options</see> for the export</param>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync(HtmlTableExportOptions options)
        {
            using (var ms = new MemoryStream())
            {
                await RenderHtmlAsync(ms, options);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        public async Task RenderHtmlAsync(Stream stream)
        {
            await RenderHtmlAsync(stream, HtmlTableExportOptions.Default);
        }

        public async Task RenderHtmlAsync(Stream stream, HtmlTableExportOptions options)
        {
            Require.Argument(options).IsNotNull("options");
            if(!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            
            var writer = new EpplusHtmlWriter(stream);
            if (_table.TableStyle != TableStyles.None)
            {
                writer.AddAttribute(HtmlAttributes.Class, $"{TableClass} {TableStyleClassPrefix}{_table.TableStyle.ToString().ToLowerInvariant()}");
            }
            else
            {
                writer.AddAttribute(HtmlAttributes.Class, $"{TableClass}");
            }
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(options.Minify);
            if (_table.ShowHeader)
            {
                await RenderHeaderRowAsync(options.Minify, writer);
            }
            // table rows
            await RenderTableRowsAsync(writer, options.Minify);
            // end tag table
            await writer.RenderEndTagAsync();

        }

        private async Task RenderTableRowsAsync(EpplusHtmlWriter writer, bool formatHtml)
        {
            await writer.RenderBeginTagAsync(HtmlElements.Tbody);
            await writer.ApplyFormatIncreaseIndentAsync(formatHtml);
            var rowIndex = _table.ShowTotal ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            while (rowIndex < _table.Address._toRow)
            {
                rowIndex++;
                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(formatHtml);
                var tableRange = _table.WorkSheet.Cells[rowIndex, _table.Address._fromCol, rowIndex, _table.Address._toCol];
                foreach (var cell in tableRange)
                {
                    await writer.RenderBeginTagAsync(HtmlElements.TableHeader);
                    // TODO: apply format
                    await writer.WriteAsync(cell.Text);
                    await writer.RenderEndTagAsync();
                    await writer.ApplyFormatAsync(formatHtml);

                }
                // end tag tr
                writer.Indent--;
                await writer.RenderEndTagAsync();
            }

            await writer.ApplyFormatDecreaseIndentAsync(formatHtml);
            // end tag tbody
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(formatHtml);
        }

        private async Task RenderHeaderRowAsync(bool formatHtml, EpplusHtmlWriter writer)
        {
            // table header row
            var rowIndex = _table.Address._fromRow;
            await writer.RenderBeginTagAsync(HtmlElements.Thead);
            await writer.ApplyFormatIncreaseIndentAsync(formatHtml);
            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(formatHtml);
            var headerRange = _table.WorkSheet.Cells[rowIndex, _table.Address._fromCol, rowIndex, _table.Address._toCol];
            foreach (var cell in headerRange)
            {
                await writer.RenderBeginTagAsync(HtmlElements.TableHeader);
                // TODO: apply format
                await writer.WriteAsync(cell.Text);
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(formatHtml);
            }
            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(formatHtml);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(formatHtml);
        }
#endif
    }
}
