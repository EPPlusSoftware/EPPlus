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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class TableExporter
    {
        internal TableExporter(ExcelTable table)
        {
            Require.Argument(table).IsNotNull("table");
            _table = table;
        }

        private readonly ExcelTable _table;
        internal const string TableClass = "epplus-table";
        internal const string TableStyleClassPrefix = "ts-";
        private readonly CellDataWriter _cellDataWriter = new CellDataWriter();
        internal List<string> _datatypes = new List<string>();
        private Dictionary<string, string> _genericCssElements;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            return GetHtmlString(HtmlTableExportOptions.Default);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="optionsConfig">Lambda for configuring <see cref="HtmlTableExportOptions">Options</see> for the export</param>
        /// <returns></returns>
        public string GetHtmlString(Action<HtmlTableExportOptions> optionsConfig)
        {
            var options = HtmlTableExportOptions.Default;
            optionsConfig.Invoke(options);
            return GetHtmlString(options);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="options"><see cref="HtmlTableExportOptions">Options</see> for the export</param>
        /// <returns>A html table</returns>
        public string GetHtmlString(HtmlTableExportOptions options)
        {
            using(var ms = new MemoryStream())
            {
                RenderHtml(ms, options);
                ms.Position = 0;
                using(var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        public void RenderHtml(Stream stream)
        {
            RenderHtml(stream, HtmlTableExportOptions.Default);
        }

        public void RenderHtml(Stream stream, HtmlTableExportOptions options)
        {
            Require.Argument(options).IsNotNull("options");
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            GetDataTypes(_table.Address);

            var writer = new EpplusHtmlWriter(stream);

            if (_table.TableStyle == TableStyles.None)
            {
                writer.AddAttribute(HtmlAttributes.Class, $"{TableClass}");
            }
            else
            {
                string styleClass;
                if (_table.TableStyle == TableStyles.Custom)
                {
                    styleClass = TableStyleClassPrefix + _table.StyleName.ToLowerInvariant();
                }
                else
                {
                    styleClass = TableStyleClassPrefix + _table.TableStyle.ToString().ToLowerInvariant();
                }

                var tblClasses = $"{TableClass} {styleClass}"; ;
                if(_table.ShowHeader)
                {
                    tblClasses += $" {styleClass}-header";
                }

                if (_table.ShowTotal)
                {
                    tblClasses += $" {styleClass}-total";
                }

                if (_table.ShowRowStripes)
                {
                    tblClasses += $" {styleClass}-row-stripes";
                }

                if (_table.ShowColumnStripes)
                {
                    tblClasses += $" {styleClass}-column-stripes";
                }

                if (_table.ShowFirstColumn)
                {
                    tblClasses += $" {styleClass}-first-column";
                }

                if (_table.ShowLastColumn)
                {
                    tblClasses += $" {styleClass}-last-column";
                }
                if(options.AdditionalTableClassNames.Count > 0)
                {
                    foreach(var cls in options.AdditionalTableClassNames)
                    {
                        tblClasses += $" {cls}";
                    }
                }

                writer.AddAttribute(HtmlAttributes.Class, tblClasses);
            }
            if(!string.IsNullOrEmpty(options.TableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, options.TableId);
            }
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(options.Minify);
            if (_table.ShowHeader)
            {
                RenderHeaderRow(writer, options);
            }
            // table rows
            RenderTableRows(writer, options);
            if(_table.ShowTotal)
            {
                RenderTotalRow(writer, options);
            }
            // end tag table
            writer.RenderEndTag();

        }
        internal string GetSinglePage(string htmlDocument = "<html><head><style>{1}</style></head><body>{0}</body></html>")
        {
            return GetSinglePage(HtmlTableExportOptions.Default, CssTableExportOptions.Default, htmlDocument);
        }

        internal string GetSinglePage(HtmlTableExportOptions htmlOptions,
                                    CssTableExportOptions cssOptions,
                                    string htmlDocument = "<html><head><style>{1}</style></head><body>{0}</body></html>")
        {
            var html = GetHtmlString(htmlOptions);
            var css = GetCssString(cssOptions);
            return string.Format(htmlDocument, html, css);

        }

        private void RenderTableRows(EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            var endRow = _table.ShowTotal ? _table.Address._toRow - 1 : _table.Address._toRow;
            while (row <= endRow)
            {
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(options.Minify);
                //var tableRange = _table.WorkSheet.Cells[row, _table.Address._fromCol, row, _table.Address._toCol];
                for (int col = _table.Address._fromCol; col <= _table.Address._toCol; col++)
                {
                    var colIx = col - _table.Address._fromCol;
                    var dataType = _datatypes[colIx];
                    _cellDataWriter.Write(_table.WorkSheet.Cells[row, col], dataType, writer, options);
                }
                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
                row++;
            }

            writer.ApplyFormatDecreaseIndent(options.Minify);
            // end tag tbody
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(options.Minify);
        }

        private void RenderHeaderRow(EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            // table header row
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            var adr = _table.Address;
            var row = adr._fromRow;
            for (int col = adr._fromCol;col <= adr._toCol; col++)
            {
                var cell = _table.WorkSheet.Cells[row, col];
                writer.AddAttribute("data-datatype", _datatypes[col - adr._fromCol]);
                writer.SetClassAttributeFromStyle(cell.StyleID, _table.WorkSheet.Workbook.Styles);
                writer.RenderBeginTag(HtmlElements.TableHeader);
                writer.Write(cell.Text);
                writer.RenderEndTag();
                writer.ApplyFormat(options.Minify);
            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(options.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(options.Minify);
        }

        private void GetDataTypes(ExcelAddressBase adr)
        {
            _datatypes = new List<string>();
            for (int col = adr._fromCol; col <= adr._toCol; col++)
            {
                _datatypes.Add(
                    ColumnDataTypeManager.GetColumnDataType(_table.WorkSheet, _table.Range, 2, col));
            }
        }
        private void RenderTotalRow(EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            // table header row
            var rowIndex = _table.Address._toRow;
            writer.RenderBeginTag(HtmlElements.TFoot);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            var address = _table.Address;
            for (var col= address._fromCol;col<= address._toCol;col++)
            {
                var cell = _table.WorkSheet.Cells[rowIndex, col];
                writer.RenderBeginTag(HtmlElements.TableData);
                writer.SetClassAttributeFromStyle(cell.StyleID, cell.Worksheet.Workbook.Styles);
                writer.Write(cell.Text);
                writer.RenderEndTag();
                writer.ApplyFormat(options.Minify);
            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(options.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(options.Minify);
        }
    }
}
