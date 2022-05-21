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
using OfficeOpenXml.Core;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlRangeExporterSync : HtmlRangeExporterSyncBase
    {
        internal HtmlRangeExporterSync
            (HtmlRangeExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
            _settings = settings;
        }

        internal HtmlRangeExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges)
        {
            _settings = settings;
        }

        private readonly HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, 0);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex)
        {
            ValidateRangeIndex(rangeIndex);
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, rangeIndex);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="settings">Override some of the settings for this html exclusively</param>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex, ExcelHtmlOverrideExportSettings settings)
        {
            ValidateRangeIndex(rangeIndex);
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, rangeIndex, settings);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public string GetHtmlString(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            return GetHtmlString(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream)
        {
            RenderHtml(stream, 0);
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="overrideSettings">Settings for this specific range index</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream, int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            ValidateRangeIndex(rangeIndex);

            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            _mergedCells.Clear();
            var range = _ranges[rangeIndex];
            GetDataTypes(range, _settings);

            ExcelTable table = null;
            if (Settings.TableStyle != eHtmlRangeTableInclude.Exclude)
            {
                table = range.GetTable();
            }

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding, _styleCache);
            var tableId = GetTableId(rangeIndex, overrideSettings);
            var additionalClassNames = GetAdditionalClassNames(overrideSettings);
            var accessibilitySettings = GetAccessibilitySettings(overrideSettings);
            var headerRows = overrideSettings != null ? overrideSettings.HeaderRows : _settings.HeaderRows;
            var headers = overrideSettings != null ? overrideSettings.Headers : _settings.Headers;
            AddClassesAttributes(writer, table, tableId, additionalClassNames);
            AddTableAccessibilityAttributes(accessibilitySettings, writer);
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            LoadVisibleColumns(range);
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(writer, range, Settings, IsMultiSheet);
            }

            if (_settings.HeaderRows > 0 || _settings.Headers.Count > 0)
            {
                RenderHeaderRow(range, writer, table, accessibilitySettings, headerRows, headers);
            }
            // table rows
            RenderTableRows(range, writer, table, accessibilitySettings);

            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            // end tag table
            writer.RenderEndTag();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public void RenderHtml(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            RenderHtml(stream, rangeIndex, settings);
        }

        /// <summary>
        /// The ranges used in the export.
        /// </summary>
        public EPPlusReadOnlyList<ExcelRangeBase> Ranges
        {
            get
            {
                return _ranges;
            }
        }

        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}\r\n</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var exporter = HtmlExporterFactory.CreateCssExporterSync(_settings, _ranges, _styleCache);
            var css = exporter.GetCssString();
            return string.Format(htmlDocument, html, css);
        }

        private void RenderTableRows(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings)
        {
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var row = range._fromRow + _settings.HeaderRows;
            var endRow = range._toRow;
            var ws = range.Worksheet;
            HtmlImage image = null;
            bool hasFooter = table != null && table.ShowTotal;
            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, range.Worksheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (hasFooter && row == endRow)
                {
                    writer.RenderBeginTag(HtmlElements.TFoot);
                }

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    writer.AddAttribute("scope", "row");
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix, IsMultiSheet);
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var colIx = col - range._fromCol;
                    var cell = ws.Cells[row, col];
                    var cv = cell.Value;
                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    SetColRowSpan(range, writer, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        _cellDataWriter.Write(cell, dataType, writer, Settings, accessibilitySettings, false, image);
                    }
                    else
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, false, Settings, imageCellClassName);
                        writer.RenderBeginTag(HtmlElements.TableData);
                        AddImage(writer, Settings, image, cell.Value);
                        RenderHyperlink(writer, cell, Settings);
                        writer.RenderEndTag();
                        writer.ApplyFormat(Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
                writer.ApplyFormat(Settings.Minify);
                if (hasFooter && row == endRow)
                {
                    writer.RenderEndTag();
                }
                row++;
            }

            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            // end tag tbody
            writer.RenderEndTag();
        }
        private void RenderHeaderRow(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings, int headerRows, List<string> headers)
        {
            if (table != null && table.ShowHeader == false) return;
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            if (table == null)
            {
                headerRows = _settings.HeaderRows == 0 ? 1 : _settings.HeaderRows;
            }
            else
            {
                headerRows = table.ShowHeader ? 1 : 0;
            }

            HtmlImage image = null;
            for (int i = 0; i < headerRows; i++)
            {
                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                }
                var row = range._fromRow + i;
                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix, IsMultiSheet);
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var cell = range.Worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        writer.AddAttribute("data-datatype", _dataTypes[col - range._fromCol]);
                    }
                    SetColRowSpan(range, writer, cell);
                    if (Settings.IncludeCssClassNames)
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, true, Settings, imageCellClassName);
                    }
                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }
                    writer.RenderBeginTag(HtmlElements.TableHeader);
                    AddImage(writer, Settings, image, cell.Value);

                    if (headerRows > 0 || table != null)
                    {
                        if (cell.Hyperlink == null)
                        {
                            writer.Write(GetCellText(cell, Settings));
                        }
                        else
                        {
                            RenderHyperlink(writer, cell, Settings);
                        }
                    }
                    else if (headers.Count < col)
                    {
                        writer.Write(headers[col]);
                    }

                    writer.RenderEndTag();
                    writer.ApplyFormat(Settings.Minify);
                }
                writer.Indent--;
                writer.RenderEndTag();
            }
            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }
    }
}
