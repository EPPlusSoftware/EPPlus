/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Table;
using System.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.Dxf;
using static OfficeOpenXml.Export.HtmlExport.ColumnDataTypeManager;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class EpplusTableCssWriter : CssWriterBase
    {
        readonly Stream _stream;
        readonly StreamWriter _writer;
        readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();
        const string IndentWhiteSpace = "  ";
        private bool _newLine;
        ExcelTable _table;
        ExcelTheme _theme;
        public EpplusTableCssWriter(Stream stream, ExcelTable table, CssTableExportOptions options)
        {
            _stream = stream;
            _table = table;
            _options = options;
            if(table.WorkSheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                table.WorkSheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = table.WorkSheet.Workbook.ThemeManager.CurrentTheme;
            _writer = new StreamWriter(stream);
        }
        internal void RenderAdditionalAndFontCss()
        {
            _writer.Write($"table.{TableExporter.TableClass}");
            _writer.Write("{");
            var ns = _table.WorkSheet.Workbook.Styles.GetNormalStyle();
            if (ns != null)
            {
                _writer.Write($"font-family:{ns.Style.Font.Name};");
                _writer.Write($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;");
            }

            foreach (var item in _options.AdditionalCssElements)
            {
                _writer.Write($"{item.Key}:{item.Value};");
            }
            _writer.Write("}");
        }
        internal void RenderCellCss(List<string> datatypes)
        {
            var styleWriter = new EpplusCssWriter(_writer, _table.Range, _options);
            styleWriter.RenderCss(datatypes);
        }
        internal void RenderTableCss(List<string> datatypes) 
        {
            ExcelTableNamedStyle tblStyle;
            if (_table.TableStyle == TableStyles.Custom)
            {
                tblStyle = _table.WorkSheet.Workbook.Styles.TableStyles[_table.StyleName].As.TableStyle;
            }
            else
            {
                var tmpNode = _table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tblStyle = new ExcelTableNamedStyle(_table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, _table.WorkSheet.Workbook.Styles);
                tblStyle.SetFromTemplate(_table.TableStyle);
            }

            var tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}";
            AddAlignmentToCss($"{tableClass}", datatypes);
            
            AddToCss($"{tableClass}", tblStyle.WholeTable, "");
            AddToCssBorderVH($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            AddToCss($"{tableClass}", tblStyle.HeaderRow, " thead tr th");
            AddToCssBorderVH($"{tableClass}", tblStyle.HeaderRow, "");

            AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            AddToCss($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            AddToCss($"{tableClass}", tblStyle.TotalRow, " tfoot tr td");
            AddToCssBorderVH($"{tableClass}", tblStyle.TotalRow, "");
            AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            AddToCss($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-column-stripes";
            AddToCss($"{tableClass}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            AddToCss($"{tableClass}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-row-stripes";
            AddToCss($"{tableClass}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            AddToCss($"{tableClass}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-last-column";
            AddToCss($"{tableClass}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-first-column";
            AddToCss($"{tableClass}", tblStyle.FirstColumn, " tbody tr td:first-child");

            _writer.Flush();
        }

        private void AddAlignmentToCss(string name, List<string> dataTypes)
        {
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            for (int c=0;c < _table.Columns.Count;c++)
            {
                var col = _table.Address._fromCol + c;
                var styleId = _table.WorkSheet.GetStyleInner(row, col);
                string hAlign = "";
                string vAlign = "";
                if(styleId>0)
                {
                    var xfs = _table.WorkSheet.Workbook.Styles.CellXfs[styleId];
                    if(xfs.ApplyAlignment)
                    {
                        hAlign = GetHorizontalAlignment(xfs);
                        vAlign = GetVerticalAlignment(xfs);
                    }
                }

                if (string.IsNullOrEmpty(hAlign) && c < dataTypes.Count && dataTypes[c] == HtmlDataTypes.Number)
                {
                    hAlign = "right";
                }

                if (!(string.IsNullOrEmpty(hAlign) && string.IsNullOrEmpty(vAlign)))
                {                    
                    _writer.Write($"table.{name} td:nth-child({col})");
                    _writer.Write("{");
                    if (string.IsNullOrEmpty(hAlign)==false && _options.Exclude.TableStyle.HorizontalAlignment==false)
                    {
                        _writer.Write($"text-align:{hAlign};");
                    }
                    if (string.IsNullOrEmpty(vAlign) == false && _options.Exclude.TableStyle.VerticalAlignment==false)
                    {
                        _writer.Write($"vertical-align:{vAlign};");
                    }
                    _writer.Write("}");
                }
            }
        }
        private void AddToCss(string name, ExcelTableStyleElement element, string htmlElement/*, bool writeFill = true, bool writeFont = true, bool writeBorder=true*/)
        {
            var s = element.Style;
            if (s.HasValue == false) return; //Dont add empty elements
            _writer.Write($"table.{name}{htmlElement}");
            _writer.Write("{");
            WriteFillStyles(s.Fill);
            WriteFontStyles(s.Font);
            WriteBorderStyles(s.Border);
            _writer.Write("}");
        }
        private void AddToCssBorderVH(string name, ExcelTableStyleElement element, string htmlElement)
        {
            var s = element.Style;
            if (s.Border.Vertical.HasValue == false && s.Border.Horizontal.HasValue==false) return; //Dont add empty elements
            _writer.Write($"table.{name}{htmlElement} td,tr");
            _writer.Write("{");
            WriteBorderStylesVerticalHorizontal(s.Border);
            _writer.Write("}");
        }
        private void WriteFillStyles(ExcelDxfFill f)
        {
            if (f.HasValue && _options.Exclude.TableStyle.Fill == false)
            {
                if (f.Style == eDxfFillStyle.PatternFill)
                {
                    if (f.PatternType.Value==ExcelFillStyle.Solid)
                    {
                        _writer.Write($"background-color:{GetDxfColor(f.BackgroundColor)};");
                    }
                    else
                    {
                        _writer.Write($"{PatternFills.GetPatternSvg(f.PatternType.Value, GetDxfColor(f.BackgroundColor), GetDxfColor(f.PatternColor))};");
                    }
                }
                else if(f.Style==eDxfFillStyle.GradientFill)
                {
                    WriteDxfGradient(f.Gradient);
                }
            }
        }

        private void WriteDxfGradient(ExcelDxfGradientFill gradient)
        {
            if(gradient.GradientType==eDxfGradientFillType.Linear)
            {
                _writer.Write($"background: linear-gradient({(gradient.Degree+90)%360}deg");
            }
            else 
            {
                _writer.Write($"background:radial-gradient(ellipse {(gradient.Right??0)*100}% {(gradient.Bottom ?? 0) * 100}%");
            }
            foreach (var color in gradient.Colors)
            {
                _writer.Write($",{GetDxfColor(color.Color)} {color.Position.ToString("F", CultureInfo.InvariantCulture)}%");
            }
            _writer.Write(")");
        }
        private void WriteFontStyles(ExcelDxfFontBase f)
        {
            var flags = _options.Exclude.TableStyle.Font;
            if (f.Color.HasValue && EnumUtil.HasNotFlag(flags, eFontExclude.Color))
            {
                _writer.Write($"color:{GetDxfColor(f.Color)};");
            }
            if (f.Bold.HasValue && f.Bold.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Bold))
            {
                _writer.Write("font-weight:bolder;");
            }
            if (f.Italic.HasValue && f.Italic.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Italic))
            {
                _writer.Write("font-style:italic;");
            }
            if (f.Strike.HasValue && f.Strike.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Strike))
            {
                _writer.Write("text-decoration:line-through solid;");
            }
            if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None && EnumUtil.HasNotFlag(flags, eFontExclude.Underline))
            {
                _writer.Write("text-decoration:underline ");
                switch (f.Underline.Value)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        _writer.Write("double;");
                        break;
                    default:
                        _writer.Write("solid;");
                        break;
                }
            }
        }
        private void WriteBorderStyles(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                var flags = _options.Exclude.TableStyle.Border;
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Top)) WriteBorderItem(b.Top, "top");
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom)) WriteBorderItem(b.Bottom, "bottom");
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Left)) WriteBorderItem(b.Left, "left");
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Right)) WriteBorderItem(b.Right, "right");
            }
        }
        private void WriteBorderStylesVerticalHorizontal(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                var flags = _options.Exclude.TableStyle.Border;
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Top)) WriteBorderItem(b.Horizontal, "top");
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom)) WriteBorderItem(b.Horizontal, "bottom");
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Left)) WriteBorderItem(b.Vertical, "left");
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Right)) WriteBorderItem(b.Vertical, "right");
            }
        }

        private void WriteBorderItem(ExcelDxfBorderItem bi, string suffix)
        {
            if (bi.HasValue && bi.Style != ExcelBorderStyle.None)
            {
                _writer.Write(WriteBorderItemLine(bi.Style.Value, suffix));
                if (bi.Color.HasValue)
                {
                    _writer.Write($" {GetDxfColor(bi.Color)}");
                }
                _writer.Write(";");
            }
        }

        private string GetDxfColor(ExcelDxfColor c)
        {
            Color ret;
            if (c.Color.HasValue)
            {
                ret = c.Color.Value;
            }
            else if (c.Theme.HasValue)
            {
                ret = ColorConverter.GetThemeColor(_theme, c.Theme.Value);
            }
            else if (c.Index.HasValue)
            {
                ret = ExcelColor.GetIndexedColor(c.Index.Value);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (c.Tint.HasValue)
            {
                ret = ColorConverter.ApplyTint(ret, c.Tint.Value);
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
    }
}
