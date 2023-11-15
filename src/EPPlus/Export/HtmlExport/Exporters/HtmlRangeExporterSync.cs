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
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;

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

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);

            var tableId = GetTableId(rangeIndex, overrideSettings);
            var additionalClassNames = GetAdditionalClassNames(overrideSettings);
            var accessibilitySettings = GetAccessibilitySettings(overrideSettings);
            var headerRows = overrideSettings != null ? overrideSettings.HeaderRows : _settings.HeaderRows;
            var headers = overrideSettings != null ? overrideSettings.Headers : _settings.Headers;

            var htmlTable = new HTMLElement(HtmlElements.Table);

            AddClassesAttributes(htmlTable, table, tableId, additionalClassNames);
            AddTableAccessibilityAttributes(accessibilitySettings, htmlTable);

            LoadVisibleColumns(range);
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(htmlTable, range, Settings, IsMultiSheet);
            }

            if (_settings.HeaderRows > 0 || _settings.Headers.Count > 0)
            {
                RenderHeaderRow(range, htmlTable, table, headers);
            }
            // table rows
            RenderTableRows(range, htmlTable, table, accessibilitySettings);

            writer.RenderHTMLElement(htmlTable, Settings.Minify);
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
            var exporter = HtmlExporterFactory.CreateCssExporterSync(_settings, _ranges, _exporterContext);
            var css = exporter.GetCssString();
            return string.Format(htmlDocument, html, css);
        }

        private void RenderTableRows(ExcelRangeBase range, HTMLElement element, ExcelTable table, AccessibilitySettings accessibilitySettings)
        {
            var tBody = new HTMLElement(HtmlElements.Tbody);
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                tBody.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
            }

            var row = range._fromRow + _settings.HeaderRows;
            var endRow = range._toRow;
            var ws = range.Worksheet;
            HtmlImage image = null;
            bool hasFooter = table != null && table.ShowTotal;
            while (row <= endRow)
            {
                EpplusHtmlAttribute attribute = null;
                if (HandleHiddenRow(attribute, range.Worksheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                HTMLElement tFoot = null;
                if (hasFooter && row == endRow)
                {
                    tFoot = new HTMLElement(HtmlElements.TFoot);
                    if(attribute != null) { tFoot.AddAttribute(attribute.AttributeName, attribute.Value); }
                    attribute = null;
                }

                var tr = new HTMLElement(HtmlElements.TableRow);

                if (attribute != null) { tr.AddAttribute(attribute.AttributeName, attribute.Value); }

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    tr.AddAttribute("role", "row");
                    tr.AddAttribute("scope", "row");
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(tr, range, row, Settings.StyleClassPrefix, IsMultiSheet);

                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var colIx = col - range._fromCol;
                    var cell = ws.Cells[row, col];
                    var cv = cell.Value;
                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    var tblData = new HTMLElement(HtmlElements.TableData);

                    SetColRowSpan(range, tblData, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        _cellDataWriter.Write(cell, dataType, tblData, Settings, accessibilitySettings, false, image, _exporterContext);
                    }
                    else
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
    
                        var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _exporterContext);

                        if (!string.IsNullOrEmpty(classString))
                        {
                            tblData.AddAttribute("class", classString);
                        }

                        AddImage(tblData, Settings, image, cell.Value);
                        AddHyperlink(tblData, cell, Settings);
                    }
                    tr.AddChildElement(tblData);
                }

                tBody.AddChildElement(tr);

                if (tFoot != null)
                {
                    tBody.AddChildElement(tFoot);
                }
                row++;
            }

            element.AddChildElement(tBody);
            //writer.RenderHTMLElement(tBody, Settings.Minify);
        }

        private void RenderHeaderRow(ExcelRangeBase range, HTMLElement element, ExcelTable table, List<string> headers)
        {
            if (table != null && table.ShowHeader == false) return;

            var thead = GetTheadAlt(range, headers);
            //var thead = GetThead(range, table, accessibilitySettings, headers);

            element.AddChildElement(thead);
        }

        //private void RenderHeaderRow(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings, List<string> headers)
        //{
        //    if (table != null && table.ShowHeader == false) return;

        //    var thead = GetThead(range, table, accessibilitySettings, headers);

        //    writer.RenderHTMLElement(thead, Settings.Minify);
        //}

        protected override int GetHeaderRows(ExcelTable table)
        {
            int headerRows;

            if (table == null)
            {
                headerRows = _settings.HeaderRows == 0 ? 1 : _settings.HeaderRows;
            }
            else
            {
                headerRows = table.ShowHeader ? 1 : 0;
            }

            return headerRows;
        }

        HTMLElement GetThead(ExcelRangeBase range, ExcelTable table, AccessibilitySettings accessibilitySettings, List<string> headers)
        {
            var thead = new HTMLElement(HtmlElements.Thead);

            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TheadRole))
            {
                thead.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            //- only here
            int headerRows;

            if (table == null)
            {
                headerRows = _settings.HeaderRows == 0 ? 1 : _settings.HeaderRows;
            }
            else
            {
                headerRows = table.ShowHeader ? 1 : 0;
            }
            //
            HtmlImage image = null;
            //- only here
            for (int i = 0; i < headerRows; i++)
            {
                //
                var tr = new HTMLElement(HtmlElements.TableRow);
                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    tr.AddAttribute("role", "row");
                }
                var row = range._fromRow + i;

                if (Settings.SetRowHeight) AddRowHeightStyle(tr, range, row, Settings.StyleClassPrefix, IsMultiSheet);

                foreach (var col in _columns)
                {
                    var th = new HTMLElement(HtmlElements.TableHeader);
                    if (InMergeCellSpan(row, col)) continue;
                    var cell = range.Worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        th.AddAttribute("data-datatype", _dataTypes[col - range._fromCol]);
                    }
                    SetColRowSpan(range, th, cell);
                    if (Settings.IncludeCssClassNames)
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, true, Settings, imageCellClassName, _exporterContext);

                        if (!string.IsNullOrEmpty(classString))
                        {
                            th.AddAttribute("class", classString);
                        }
                    }
                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    AddImage(th, Settings, image, cell.Value);

                    if (headerRows > 0 || table != null)
                    {
                        if (cell.Hyperlink == null)
                        {
                            th.Content = GetCellText(cell, Settings);
                        }
                        else
                        {
                            AddHyperlink(th, cell, Settings);
                        }
                    }
                    else if (headers.Count < col)
                    {
                        th.Content = headers[col];
                    }
                    tr.AddChildElement(th);
                }
                thead.AddChildElement(tr);
            }
            return thead;
        }
    }
}
