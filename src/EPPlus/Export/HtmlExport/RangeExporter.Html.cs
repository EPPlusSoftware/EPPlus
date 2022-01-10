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
    public partial class RangeExporter
    {        
        internal RangeExporter
            (ExcelRangeBase range)
        {
            Require.Argument(range).IsNotNull("range");
            _range = range;
        }

        private readonly ExcelRangeBase _range;
        internal const string TableClass = "range";
        internal const string TableStyleClassPrefix = "s-";
        private readonly CellDataWriter _cellDataWriter = new CellDataWriter();
        internal List<string> _datatypes = new List<string>();
        private List<int> _columns = new List<int>();
        public HtmlRangeExportSettings Settings { get; } = new HtmlRangeExportSettings();
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using (var ms = new MemoryStream())
            {
                RenderHtml(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
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

            GetDataTypes();

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            LoadVisibleColumns();
            if (Settings.FirstRowIsHeader || Settings.Headers.Count > 0)
            {
                RenderHeaderRow(writer);
            }
            // table rows
            RenderTableRows(writer);

            // end tag table
            writer.RenderEndTag();

        }

        private void LoadVisibleColumns()
        {
            var ws = _range.Worksheet;
            _columns = new List<int>();
            for (int col = _range._fromCol; col <= _range._toCol; col++)
            {
                var c = ws.GetColumn(col);
                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    _columns.Add(col);
                }
            }
        }


        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var css = GetCssString();
            return string.Format(htmlDocument, html, css);
            //return string.Format(htmlDocument, html, "");

        }

        private void RenderTableRows(EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var row = Settings.FirstRowIsHeader ? _range._fromRow + 1 : _range._fromRow;
            var endRow = _range._toRow;
            var ws = _range.Worksheet;
            while (row <= endRow)
            {
                if (Settings.IncludeHiddenRows)
                {
                    var r = ws.Row(row);
                    if (r.Hidden || r.Height == 0)
                    {
                        continue;
                    }

                    if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                    {
                        writer.AddAttribute("role", "row");
                        writer.AddAttribute("scope", "row");
                    }
                }

                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);
                foreach (var col in _columns)
                {
                    var colIx = col - _range._fromCol;
                    var dataType = _datatypes[colIx];
                    var cell = ws.Cells[row, col];

                    if (cell.Hyperlink == null)
                    {
                        _cellDataWriter.Write(cell, dataType, writer, Settings, false);
                    }
                    else
                    {
                        writer.RenderBeginTag(HtmlElements.TableData);
                        SetColRowSpan(writer, cell);
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
            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var row = _range._fromRow;
            foreach (var col in _columns)
            {
                var cell = _range.Worksheet.Cells[row, col];
                writer.AddAttribute("data-datatype", _datatypes[col - _range._fromCol]);
                SetColRowSpan(writer, cell);
                writer.SetClassAttributeFromStyle(cell.StyleID, _range.Worksheet.Workbook.Styles);
                writer.RenderBeginTag(HtmlElements.TableHeader);
                if (Settings.FirstRowIsHeader)
                {
                    if (cell.Hyperlink == null)
                    {
                        writer.Write(GetCellText(cell));
                    }
                    else
                    {
                        RenderHyperlink(writer, cell);
                    }
                }
                else if (Settings.Headers.Count < col)
                {
                    writer.Write(Settings.Headers[col]);
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

        private void SetColRowSpan(EpplusHtmlWriter writer, ExcelRange cell)
        {
            if(cell.Merge)
            {
                var address = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
                if(address!=null)
                {
                    var ma = new ExcelAddressBase(address);
                    //ColSpan
                    if(ma._fromCol==cell._fromCol || _range._fromCol==cell._fromCol)
                    {
                        var maxCol = Math.Min(cell._toCol, _range._toCol);
                        var colSpan = maxCol - ma._fromCol;
                        if(colSpan>1)
                        {
                            writer.AddAttribute("colspan", colSpan.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                    //RowSpan
                    if (ma._fromRow == cell._fromRow || _range._fromRow == cell._fromRow)
                    {
                        var maxRow = Math.Min(cell._toRow, _range._toRow);
                        var rowSpan = maxRow - ma._fromRow;
                        if (rowSpan > 1)
                        {
                            writer.AddAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                }
            }
        }

        private void RenderHyperlink(EpplusHtmlWriter writer, ExcelRangeBase cell)
        {
            if (cell.Hyperlink is ExcelHyperLink eurl)
            {
                if (string.IsNullOrEmpty(eurl.ReferenceAddress))
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
            if (cell.IsRichText)
            {
                return cell.RichText.HtmlText;
            }
            else
            {
                return ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, Settings.Culture);
            }
        }

        private void GetDataTypes()
        {
            _datatypes = new List<string>();
            for (int col = _range._fromCol; col <= _range._toCol; col++)
            {
                _datatypes.Add(
                    ColumnDataTypeManager.GetColumnDataType(_range.Worksheet, _range, Settings.FirstRowIsHeader ? 2 : 1, col));
            }
        }
    }
}

