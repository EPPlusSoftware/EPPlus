/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlTableExporterAsync : HtmlRangeExporterAsyncBase
    {
        public HtmlTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table) : base(settings, table.Range)
        {
            Require.Argument(table).IsNotNull("table");
            _table = table;
            _settings = settings;
            LoadRangeImages(new List<ExcelRangeBase>() { table.Range });
        }

        private readonly ExcelTable _table;
        private HtmlTableExportSettings _settings;

        private void LoadVisibleColumns()
        {
            _columns = new List<int>();
            var r = _table.Range;
            for (int col = r._fromCol; col <= r._toCol; col++)
            {
                var c = _table.WorkSheet.GetColumn(col);
                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    _columns.Add(col);
                }
            }
        }

        private async Task RenderTableRowsAsync(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
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

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    if (!_table.ShowFirstColumn && !_table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "row");
                    }
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, row, Settings.StyleClassPrefix, false);

                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
                foreach (var col in _columns)
                {
                    var colIx = col - _table.Address._fromCol;
                    var dataType = _dataTypes[colIx];
                    var cell = _table.WorkSheet.Cells[row, col];

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        var addRowScope = (_table.ShowFirstColumn && col == _table.Address._fromCol) || (_table.ShowLastColumn && col == _table.Address._toCol);
                        await _cellDataWriter.WriteAsync(cell, dataType, writer, Settings, accessibilitySettings, addRowScope, image, _cfAtAddresses);
                    }
                    else
                    {
                        await writer.RenderBeginTagAsync(HtmlElements.TableData);
                        var imageCellClassName = GetImageCellClassName(image, Settings);

                        var classString = AttributeParser.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _cfAtAddresses, writer._styleCache, writer._dxfStyleCache);

                        if (!string.IsNullOrEmpty(classString))
                        {
                            writer.AddAttribute("class", classString);
                        }

                        await RenderHyperlinkAsync(writer, cell, Settings);
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


        private async Task RenderHeaderRowAsync(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            // table header row
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TheadRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Thead);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
            }
            var adr = _table.Address;
            var row = adr._fromRow;
            if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, row, Settings.StyleClassPrefix, false);
            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var cell = _table.WorkSheet.Cells[row, col];
                if (Settings.RenderDataTypes)
                {
                    writer.AddAttribute("data-datatype", _dataTypes[col - adr._fromCol]);
                }

                var imageCellClassName = image == null ? "" : Settings.StyleClassPrefix + "image-cell";

                var classString = AttributeParser.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _cfAtAddresses, writer._styleCache, writer._dxfStyleCache);

                if (!string.IsNullOrEmpty(classString))
                {
                    writer.AddAttribute("class", classString);
                }

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TableHeaderCellRole))
                {
                    writer.AddAttribute("role", accessibilitySettings.TableSettings.TableHeaderCellRole);
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
                    image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                }
                await AddImageAsync(writer, Settings, image, cell.Value);

                if (cell.Hyperlink == null)
                {
                    await writer.WriteAsync(GetCellText(cell, Settings));
                }
                else
                {
                    await RenderHyperlinkAsync(writer, cell, Settings);
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
        private async Task RenderTotalRowAsync(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            // table header row
            var rowIndex = _table.Address._toRow;
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TfootRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TfootRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.TFoot);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
                writer.AddAttribute("scope", "row");
            }
            if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, rowIndex, Settings.StyleClassPrefix, false);
            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            var address = _table.Address;
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var cell = _table.WorkSheet.Cells[rowIndex, col];
                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "cell");
                }
                var imageCellClassName = GetImageCellClassName(image, Settings);
                var classString = AttributeParser.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _cfAtAddresses, writer._styleCache, writer._dxfStyleCache);

                if (!string.IsNullOrEmpty(classString))
                {
                    writer.AddAttribute("class", classString);
                }

                await writer.RenderBeginTagAsync(HtmlElements.TableData);
                await AddImageAsync(writer, Settings, image, cell.Value);
                await writer.WriteAsync(GetCellText(cell, Settings));
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(Settings.Minify);
            }
            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }

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
            GetDataTypes(_table.Address, _table);

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding, _exporterContext._dxfStyleCache);
            HtmlExportTableUtil.AddClassesAttributes(writer, _table, _settings);
            AddTableAccessibilityAttributes(Settings.Accessibility, writer);
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            LoadVisibleColumns();
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                await SetColumnGroupAsync(writer, _table.Range, Settings, false);
            }

            if (_table.ShowHeader)
            {
                await RenderHeaderRowAsync(writer, Settings.Accessibility);
            }
            // table rows
            await RenderTableRowsAsync(writer, Settings.Accessibility);
            if (_table.ShowTotal)
            {
                await RenderTotalRowAsync(writer, Settings.Accessibility);
            }
            // end tag table
            await writer.RenderEndTagAsync();

        }
        /// <summary>
        /// Renders the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = await GetHtmlStringAsync();
            var cssExporter = HtmlExporterFactory.CreateCssExporterTableAsync(_settings, _table, _exporterContext);
            var css = await cssExporter.GetCssStringAsync();
            return string.Format(htmlDocument, html, css);
        }
    }
}
#endif
