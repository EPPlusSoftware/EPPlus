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
using OfficeOpenXml.Export.HtmlExport.Accessibility;
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
            AddTableAccessibilityAttributes(options, writer);
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

        private void AddTableAccessibilityAttributes(HtmlTableExportOptions options, EpplusHtmlWriter writer)
        {
            if (!options.Accessibility.TableSettings.AddAccessibilityAttributes) return;
            if(!string.IsNullOrEmpty(options.Accessibility.TableSettings.TableRole))
            {
                writer.AddAttribute("role", options.Accessibility.TableSettings.TableRole);
            }
            if(!string.IsNullOrEmpty(options.Accessibility.TableSettings.AriaLabel))
            {
                writer.AddAttribute(AriaAttributes.AriaLabel.AttributeName, options.Accessibility.TableSettings.AriaLabel);
            }
            if (!string.IsNullOrEmpty(options.Accessibility.TableSettings.AriaLabelledBy))
            {
                writer.AddAttribute(AriaAttributes.AriaLabelledBy.AttributeName, options.Accessibility.TableSettings.AriaLabelledBy);
            }
            if (!string.IsNullOrEmpty(options.Accessibility.TableSettings.AriaDescribedBy))
            {
                writer.AddAttribute(AriaAttributes.AriaDescribedBy.AttributeName, options.Accessibility.TableSettings.AriaDescribedBy);
            }
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
            if (options.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(options.Accessibility.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", options.Accessibility.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            var endRow = _table.ShowTotal ? _table.Address._toRow - 1 : _table.Address._toRow;
            while (row <= endRow)
            {
                if(options.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    if (!_table.ShowFirstColumn && !_table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "row");
                    }
                }
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(options.Minify);
                //var tableRange = _table.WorkSheet.Cells[row, _table.Address._fromCol, row, _table.Address._toCol];
                for (int col = _table.Address._fromCol; col <= _table.Address._toCol; col++)
                {
                    var colIx = col - _table.Address._fromCol;
                    var dataType = _datatypes[colIx];
                    var addRowScope = (col == _table.Address._fromCol && _table.ShowFirstColumn) || (col == _table.Address._toCol && _table.ShowLastColumn);
                    _cellDataWriter.Write(_table.WorkSheet.Cells[row, col], dataType, writer, options, addRowScope);
                }
                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
                writer.ApplyFormat(options.Minify);
                row++;
            }

            writer.ApplyFormatDecreaseIndent(options.Minify);
            // end tag tbody
            writer.RenderEndTag();
            writer.ApplyFormat(options.Minify);
        }

        private void RenderHeaderRow(EpplusHtmlWriter writer, HtmlTableExportOptions options)
        {
            // table header row
            if(options.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(options.Accessibility.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", options.Accessibility.TableSettings.TheadRole);
            }
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            if (options.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
            }
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            var adr = _table.Address;
            var row = adr._fromRow;
            for (int col = adr._fromCol;col <= adr._toCol; col++)
            {
                var cell = _table.WorkSheet.Cells[row, col];
                writer.AddAttribute("data-datatype", _datatypes[col - adr._fromCol]);
                writer.SetClassAttributeFromStyle(cell.StyleID, _table.WorkSheet.Workbook.Styles);
                if (options.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(options.Accessibility.TableSettings.TableHeaderCellRole))
                {
                    writer.AddAttribute("role", options.Accessibility.TableSettings.TableHeaderCellRole);
                    if(!_table.ShowFirstColumn && !_table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "col");
                    }
                    if(_table.SortState != null && !_table.SortState.ColumnSort && _table.SortState.SortConditions.Any())
                    {
                        var firstCondition = _table.SortState.SortConditions.First();
                        if(firstCondition != null && !string.IsNullOrEmpty(firstCondition.Ref))
                        {
                            var addr = new ExcelAddress(firstCondition.Ref);
                            var sortedCol = addr._fromCol;
                            if(col == sortedCol)
                            {
                                writer.AddAttribute("aria-sort", firstCondition.Descending ? "descending" : "ascending");
                            }
                        }
                    }
                }
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
            if (options.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(options.Accessibility.TableSettings.TfootRole))
            {
                writer.AddAttribute("role", options.Accessibility.TableSettings.TfootRole);
            }
            writer.RenderBeginTag(HtmlElements.TFoot);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            if (options.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
                writer.AddAttribute("scope", "row");
            }
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(options.Minify);
            var address = _table.Address;
            for (var col= address._fromCol;col<= address._toCol;col++)
            {
                var cell = _table.WorkSheet.Cells[rowIndex, col];
                if(options.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "cell");
                }
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
