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
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
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

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);
            var htmlTable = new HTMLElement(HtmlElements.Table);

            HtmlExportTableUtil.AddClassesAttributes(htmlTable, _table, _settings);
            AddTableAccessibilityAttributes(Settings.Accessibility, htmlTable);

            LoadVisibleColumns();
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(htmlTable, _table.Range, Settings, false);
            }

            if (_table.ShowHeader)
            {
                AddHeaderRow(htmlTable);
            }

            // table rows
            AddTableRows(htmlTable);
            if (_table.ShowTotal)
            {
                RenderTotalRow(htmlTable);
            }
            //// end tag table
            //await writer.RenderEndTagAsync();
            await writer.RenderHTMLElementAsync(htmlTable, Settings.Minify);
        }

        void AddTableRows(HTMLElement htmlTable)
        {
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            var endRow = _table.ShowTotal ? _table.Address._toRow - 1 : _table.Address._toRow;

            var body = GetTableBody(_table.Range, row, endRow);
            htmlTable.AddChildElement(body);
        }

        private void AddHeaderRow(HTMLElement table)
        {
            table.AddChildElement(GetThead(_table.Range));
        }

        private void RenderTotalRow(HTMLElement table)
        {
            // table header row
            var tFoot = new HTMLElement(HtmlElements.TFoot);

            var rowIndex = _table.Address._toRow;
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TfootRole))
            {
                tFoot.AddAttribute("role", Settings.Accessibility.TableSettings.TfootRole);
            }

            var row = new HTMLElement(HtmlElements.TableRow);

            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                row.AddAttribute("role", "row");
                row.AddAttribute("scope", "row");
            }
            if (Settings.SetRowHeight) AddRowHeightStyle(row, _table.Range, rowIndex, Settings.StyleClassPrefix, false);

            var address = _table.Address;
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var tblData = new HTMLElement(HtmlElements.TableData);

                var cell = _table.WorkSheet.Cells[rowIndex, col];
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    tblData.AddAttribute("role", "cell");
                }
                var imageCellClassName = GetImageCellClassName(image, Settings);

                var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _exporterContext);

                if (!string.IsNullOrEmpty(classString))
                {
                    tblData.AddAttribute("class", classString);
                }

                AddImage(tblData, Settings, image, cell.Value);

                tblData.Content = GetCellText(cell, Settings);

                row.AddChildElement(tblData);
            }
            tFoot.AddChildElement(row);
            table.AddChildElement(tFoot);
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
