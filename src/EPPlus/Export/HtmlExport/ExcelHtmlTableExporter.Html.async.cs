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
using System.IO;
using System.Linq;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    public partial class ExcelHtmlTableExporter : HtmlExporterBase
    {
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderHtmlAsync(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Exports the html part of an <see cref="ExcelTable"/> to a stream
        /// </summary>
        /// <returns>A html table</returns>
        public async Task RenderHtmlAsync(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            GetDataTypes(_table.Address);

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);
            AddClassesAttributes(writer);
            AddTableAccessibilityAttributes(Settings, writer);
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            LoadVisibleColumns();
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                await SetColumnGroupAsync(writer, _table.Range, Settings);
            }

            if (_table.ShowHeader)
            {
                await RenderHeaderRowAsync(writer);
            }
            // table rows
            await RenderTableRowsAsync(writer);
            if (_table.ShowTotal)
            {
                await RenderTotalRowAsync(writer);
            }
            // end tag table
            await writer.RenderEndTagAsync();

        }
        /// <summary>
        /// Renders the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = await GetHtmlStringAsync();
            var css = await GetCssStringAsync();
            return string.Format(htmlDocument, html, css);

        }

        private async Task RenderTableRowsAsync(EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TbodyRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Tbody);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            var endRow = _table.ShowTotal ? _table.Address._toRow - 1 : _table.Address._toRow;
            HtmlImage image = null;
            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, _table.WorkSheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    if (!_table.ShowFirstColumn && !_table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "row");
                    }
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, row, Settings.StyleClassPrefix);

                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
                foreach (var col in _columns)
                {
                    var colIx = col - _table.Address._fromCol;
                    var dataType = _datatypes[colIx];
                    var cell = _table.WorkSheet.Cells[row, col];

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        var addRowScope = (_table.ShowFirstColumn && col == _table.Address._fromCol) || (_table.ShowLastColumn && col == _table.Address._toCol);
                        await _cellDataWriter.WriteAsync(cell, dataType, writer, Settings, addRowScope, image);
                    }
                    else
                    {
                        await writer.RenderBeginTagAsync(HtmlElements.TableData);
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, Settings.HorizontalAlignmentWhenGeneral, false, Settings.StyleClassPrefix, imageCellClassName);
                        await RenderHyperlinkAsync(writer, cell);
                        await writer.RenderEndTagAsync();
                        await writer.ApplyFormatAsync(Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(Settings.Minify);
                row++;
            }

            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            // end tag tbody
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }


        private async Task RenderHeaderRowAsync(EpplusHtmlWriter writer)
        {
            // table header row
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Thead);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
            }
            var adr = _table.Address;
            var row = adr._fromRow;
            if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, row, Settings.StyleClassPrefix);
            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var cell = _table.WorkSheet.Cells[row, col];
                if (Settings.RenderDataTypes)
                {
                    writer.AddAttribute("data-datatype", _datatypes[col - adr._fromCol]);
                }

                var imageCellClassName = image == null ? "" : Settings.StyleClassPrefix + "image-cell";
                writer.SetClassAttributeFromStyle(cell, Settings.HorizontalAlignmentWhenGeneral, true, Settings.StyleClassPrefix, imageCellClassName);
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TableHeaderCellRole))
                {
                    writer.AddAttribute("role", Settings.Accessibility.TableSettings.TableHeaderCellRole);
                    if (!_table.ShowFirstColumn && !_table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "col");
                    }
                    if (_table.SortState != null && !_table.SortState.ColumnSort && _table.SortState.SortConditions.Any())
                    {
                        var firstCondition = _table.SortState.SortConditions.First();
                        if (firstCondition != null && !string.IsNullOrEmpty(firstCondition.Ref))
                        {
                            var addr = new ExcelAddress(firstCondition.Ref);
                            var sortedCol = addr._fromCol;
                            if (col == sortedCol)
                            {
                                writer.AddAttribute("aria-sort", firstCondition.Descending ? "descending" : "ascending");
                            }
                        }
                    }
                }
                await writer.RenderBeginTagAsync(HtmlElements.TableHeader);
                if (Settings.Pictures.Include == ePictureInclude.Include)
                {
                    image = GetImage(cell._fromRow, cell._fromCol);
                }
                await AddImageAsync(writer, Settings, image, cell.Value);

                if (cell.Hyperlink == null)
                {
                    await writer.WriteAsync(GetCellText(cell));
                }
                else
                {
                    await RenderHyperlinkAsync(writer, cell);
                }
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(Settings.Minify);
            }
            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }
        private async Task RenderHyperlinkAsync(EpplusHtmlWriter writer, ExcelRangeBase cell)
        {
            if (cell.Hyperlink is ExcelHyperLink eurl)
            {
                if (string.IsNullOrEmpty(eurl.ReferenceAddress))
                {
                    writer.AddAttribute("href", eurl.AbsolutePath);
                    await writer.RenderBeginTagAsync(HtmlElements.A);
                    await writer.WriteAsync(eurl.Display);
                    await writer.RenderEndTagAsync();
                }
                else
                {
                    //Internal
                    await writer.WriteAsync(GetCellText(cell));
                }
            }
            else
            {
                writer.AddAttribute("href", cell.Hyperlink.OriginalString);
                await writer.RenderBeginTagAsync(HtmlElements.A);
                await writer.WriteAsync(GetCellText(cell));
                await writer.RenderEndTagAsync();
            }
        }
        private async Task RenderTotalRowAsync(EpplusHtmlWriter writer)
        {
            // table header row
            var rowIndex = _table.Address._toRow;
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TfootRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TfootRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.TFoot);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
                writer.AddAttribute("scope", "row");
            }
            if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, rowIndex, Settings.StyleClassPrefix);
            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            var address = _table.Address;
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var cell = _table.WorkSheet.Cells[rowIndex, col];
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "cell");
                }
                var imageCellClassName = GetImageCellClassName(image, Settings);
                writer.SetClassAttributeFromStyle(cell, Settings.HorizontalAlignmentWhenGeneral, false, Settings.StyleClassPrefix, imageCellClassName);
                await writer.RenderBeginTagAsync(HtmlElements.TableData);
                await AddImageAsync(writer, Settings, image, cell.Value);
                await writer.WriteAsync(GetCellText(cell));
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(Settings.Minify);
            }
            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }
    }
#endif
}

