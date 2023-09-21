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
using OfficeOpenXml.Utils;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.Parsers;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlTableExporterSync : HtmlRangeExporterSyncBase
    {
        internal HtmlTableExporterSync(HtmlTableExportSettings settings, ExcelTable table)
            : base(settings, table.Range)
        {
            Require.Argument(table).IsNotNull("table");
            _table = table;
            _tableExportSettings = settings;

            LoadRangeImages(new List<ExcelRangeBase>() { table.Range });
        }

        private readonly ExcelTable _table;
        private readonly HtmlTableExportSettings _tableExportSettings;

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

        private void RenderHeaderRow(EpplusHtmlWriter writer)
        {
            // table header row
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
            }
            var adr = _table.Address;
            var row = adr._fromRow;
            if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, row, Settings.StyleClassPrefix, false);
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var cell = _table.WorkSheet.Cells[row, col];
                if (Settings.RenderDataTypes)
                {
                    writer.AddAttribute("data-datatype", _dataTypes[col - adr._fromCol]);
                }
                var imageCellClassName = image == null ? "" : Settings.StyleClassPrefix + "image-cell";

                var classString = AttributeParser.GetClassAttributeFromStyle(cell, true, Settings, imageCellClassName, _cfAtAddresses, writer._styleCache, writer._dxfStyleCache);

                if (!string.IsNullOrEmpty(classString))
                {
                    writer.AddAttribute("class", classString);
                }

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
                writer.RenderBeginTag(HtmlElements.TableHeader);
                if (Settings.Pictures.Include == ePictureInclude.Include)
                {
                    image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                }
                AddImage(writer, Settings, image, cell.Value);

                if (cell.Hyperlink == null)
                {
                    writer.Write(GetCellText(cell, Settings));
                }
                else
                {
                    RenderHyperlink(writer, cell, Settings);
                }

                writer.RenderEndTag();
                writer.ApplyFormat(Settings.Minify);
            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }

        private void RenderTableRows(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
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

                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);

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
                        _cellDataWriter.Write(cell, dataType, writer, Settings, accessibilitySettings, addRowScope, image, _cfAtAddresses);
                    }
                    else
                    {
                        writer.RenderBeginTag(HtmlElements.TableData);
                        AddImage(writer, Settings, image, cell.Value);
                        var imageCellClassName = GetImageCellClassName(image, Settings);

                        var classString = AttributeParser.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _cfAtAddresses, writer._styleCache, writer._dxfStyleCache);

                        if (!string.IsNullOrEmpty(classString))
                        {
                            writer.AddAttribute("class", classString);
                        }
                        RenderHyperlink(writer, cell, Settings);
                        writer.RenderEndTag();
                        writer.ApplyFormat(Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
                writer.ApplyFormat(Settings.Minify);
                row++;
            }

            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            // end tag tbody
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }

        private void RenderTotalRow(EpplusHtmlWriter writer)
        {
            // table header row
            var rowIndex = _table.Address._toRow;
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TfootRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TfootRole);
            }
            writer.RenderBeginTag(HtmlElements.TFoot);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
                writer.AddAttribute("scope", "row");
            }
            if (Settings.SetRowHeight) AddRowHeightStyle(writer, _table.Range, rowIndex, Settings.StyleClassPrefix, false);
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
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

                var classString = AttributeParser.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _cfAtAddresses, writer._styleCache, writer._dxfStyleCache);

                if (!string.IsNullOrEmpty(classString))
                {
                    writer.AddAttribute("class", classString);
                }

                writer.RenderBeginTag(HtmlElements.TableData);
                AddImage(writer, Settings, image, cell.Value);
                writer.Write(GetCellText(cell, Settings));
                writer.RenderEndTag();
                writer.ApplyFormat(Settings.Minify);
            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports the html part of an <see cref="ExcelTable"/> to a html string.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <exception cref="IOException"></exception>
        public void RenderHtml(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            GetDataTypes(_table.Address, _table);

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding, _exporterContext._styleCache);
            HtmlExportTableUtil.AddClassesAttributes(writer, _table, _tableExportSettings);
            AddTableAccessibilityAttributes(Settings.Accessibility, writer);
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            LoadVisibleColumns();
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(writer, _table.Range, Settings, false);
            }

            if (_table.ShowHeader)
            {
                RenderHeaderRow(writer);
            }
            // table rows
            RenderTableRows(writer, Settings.Accessibility);
            if (_table.ShowTotal)
            {
                RenderTotalRow(writer);
            }
            // end tag table
            writer.RenderEndTag();
        }

        /// <summary>
        /// Renders both the Css and the Html to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var cssExporter = HtmlExporterFactory.CreateCssExporterTableSync(_tableExportSettings, _table, _exporterContext);
            var css = cssExporter.GetCssString();
            return string.Format(htmlDocument, html, css);

        }
    }
}
