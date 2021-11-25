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
        public HtmlTableExportSettings Settings { get; } = new HtmlTableExportSettings();
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using(var ms = new MemoryStream())
            {
                RenderHtml(ms);
                ms.Position = 0;
                using(var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        public void RenderHtml(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            GetDataTypes(_table.Address);
            
            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);

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
                if(Settings.AdditionalTableClassNames.Count > 0)
                {
                    foreach(var cls in Settings.AdditionalTableClassNames)
                    {
                        tblClasses += $" {cls}";
                    }
                }

                writer.AddAttribute(HtmlAttributes.Class, tblClasses);
            }
            if(!string.IsNullOrEmpty(Settings.TableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, Settings.TableId);
            }
            AddTableAccessibilityAttributes(Settings, writer);
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            if (_table.ShowHeader)
            {
                RenderHeaderRow(writer);
            }
            // table rows
            RenderTableRows(writer);
            if(_table.ShowTotal)
            {
                RenderTotalRow(writer);
            }
            // end tag table
            writer.RenderEndTag();

        }

        private void AddTableAccessibilityAttributes(HtmlTableExportSettings settings, EpplusHtmlWriter writer)
        {
            if (!settings.Accessibility.TableSettings.AddAccessibilityAttributes) return;
            if(!string.IsNullOrEmpty(settings.Accessibility.TableSettings.TableRole))
            {
                writer.AddAttribute("role", settings.Accessibility.TableSettings.TableRole);
            }
            if(!string.IsNullOrEmpty(settings.Accessibility.TableSettings.AriaLabel))
            {
                writer.AddAttribute(AriaAttributes.AriaLabel.AttributeName, settings.Accessibility.TableSettings.AriaLabel);
            }
            if (!string.IsNullOrEmpty(settings.Accessibility.TableSettings.AriaLabelledBy))
            {
                writer.AddAttribute(AriaAttributes.AriaLabelledBy.AttributeName, settings.Accessibility.TableSettings.AriaLabelledBy);
            }
            if (!string.IsNullOrEmpty(settings.Accessibility.TableSettings.AriaDescribedBy))
            {
                writer.AddAttribute(AriaAttributes.AriaDescribedBy.AttributeName, settings.Accessibility.TableSettings.AriaDescribedBy);
            }
        }

        /// <summary>
        /// Renders both the Css and the Html to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var css = GetCssString();
            return string.Format(htmlDocument, html, css);

        }

        private void RenderTableRows(EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            var endRow = _table.ShowTotal ? _table.Address._toRow - 1 : _table.Address._toRow;
            while (row <= endRow)
            {
                if (Settings.IncludeHiddenRows)
                {
                    var r = _table.WorkSheet.Row(row);
                    if (r.Hidden || r.Height == 0)
                    {
                        continue;
                    }

                    if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                    {
                        writer.AddAttribute("role", "row");
                        if (!_table.ShowFirstColumn && !_table.ShowLastColumn)
                        {
                            writer.AddAttribute("scope", "row");
                        }
                    }
                }

                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);
                //var tableRange = _table.WorkSheet.Cells[row, _table.Address._fromCol, row, _table.Address._toCol];
                for (int col = _table.Address._fromCol; col <= _table.Address._toCol; col++)
                {
                    var colIx = col - _table.Address._fromCol;
                    var dataType = _datatypes[colIx];
                    var cell = _table.WorkSheet.Cells[row, col];
                    
                    if (cell.Hyperlink == null)
                    {
                        var addRowScope = (_table.ShowFirstColumn && col == _table.Address._fromCol) || (_table.ShowLastColumn && col == _table.Address._toCol);
                        _cellDataWriter.Write(cell, dataType, writer, Settings, addRowScope);
                    }
                    else
                    {
                        writer.RenderBeginTag(HtmlElements.TableData);
                        writer.SetClassAttributeFromStyle(cell.StyleID, cell.Worksheet.Workbook.Styles);
                        RenderHyperlink(writer, cell);
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

        private void RenderHeaderRow(EpplusHtmlWriter writer)
        {
            // table header row
            if(Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
            }
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var adr = _table.Address;
            var row = adr._fromRow;
            for (int col = adr._fromCol;col <= adr._toCol; col++)
            {
                var cell = _table.WorkSheet.Cells[row, col];
                writer.AddAttribute("data-datatype", _datatypes[col - adr._fromCol]);
                writer.SetClassAttributeFromStyle(cell.StyleID, _table.WorkSheet.Workbook.Styles);
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TableHeaderCellRole))
                {
                    writer.AddAttribute("role", Settings.Accessibility.TableSettings.TableHeaderCellRole);
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
                if (cell.Hyperlink == null)
                {
                    writer.Write(GetCellText(cell));
                }
                else
                {
                    RenderHyperlink(writer, cell);
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

        private void RenderHyperlink(EpplusHtmlWriter writer, ExcelRangeBase cell)
        {
            if(cell.Hyperlink is ExcelHyperLink eurl)
            {
                if(string.IsNullOrEmpty(eurl.ReferenceAddress))
                {
                    writer.AddAttribute("href", eurl.AbsolutePath);
                    writer.RenderBeginTag(HtmlElements.A);
                    writer.Write(eurl.Display);
                    writer.RenderEndTag();
                }
                else
                {
                    //Internal
                    writer.Write(GetCellText(cell)); 
                }
            }
            else
            {
                writer.AddAttribute("href", cell.Hyperlink.OriginalString);
                writer.RenderBeginTag(HtmlElements.A);
                writer.Write(GetCellText(cell));
                writer.RenderEndTag();
            }
        }

        private string GetCellText(ExcelRangeBase cell)
        {
            return ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, Settings.Culture);
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
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var address = _table.Address;
            for (var col= address._fromCol;col<= address._toCol;col++)
            {
                var cell = _table.WorkSheet.Cells[rowIndex, col];
                if(Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "cell");
                }
                writer.SetClassAttributeFromStyle(cell.StyleID, cell.Worksheet.Workbook.Styles);
                writer.RenderBeginTag(HtmlElements.TableData);
                writer.Write(GetCellText(cell));
                writer.RenderEndTag();
                writer.ApplyFormat(Settings.Minify);
            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }
    }
}
