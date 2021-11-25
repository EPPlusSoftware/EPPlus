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
    internal class EpplusTableCssWriter : HtmlWriterBase
    {
        protected HtmlTableExportSettings _settings;
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();
        ExcelTable _table;
        ExcelTheme _theme;
        internal EpplusTableCssWriter(Stream stream, ExcelTable table, HtmlTableExportSettings settings) : base(stream, settings.Encoding)
        {
            Init(table, settings);
        }
        internal EpplusTableCssWriter(StreamWriter writer, ExcelTable table, HtmlTableExportSettings settings) : base(writer)
        {
            Init(table, settings);
        }
        private void Init(ExcelTable table, HtmlTableExportSettings settings)
        {
            _table = table;
            _settings = settings;
            if (table.WorkSheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                table.WorkSheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = table.WorkSheet.Workbook.ThemeManager.CurrentTheme;
        }

        internal void RenderAdditionalAndFontCss()
        {
            WriteClass($"table.{TableExporter.TableClass}{{", _settings.Minify);
            var ns = _table.WorkSheet.Workbook.Styles.GetNormalStyle();
            if (ns != null)
            {
                WriteCssItem($"font-family:{ns.Style.Font.Name};", _settings.Minify);
                WriteCssItem($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            foreach (var item in _settings.Css.AdditionalCssElements)
            {
                WriteCssItem($"{item.Key}:{item.Value};", _settings.Minify);
            }
            WriteClassEnd(_settings.Minify);
        }

        internal void AddAlignmentToCss(string name, List<string> dataTypes)
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
                    WriteClass($"table.{name} td:nth-child({col}){{", _settings.Minify);
                    if (string.IsNullOrEmpty(hAlign)==false && _settings.Css.Exclude.TableStyle.HorizontalAlignment==false)
                    {
                        WriteCssItem($"text-align:{hAlign};", _settings.Minify);
                    }
                    if (string.IsNullOrEmpty(vAlign) == false && _settings.Css.Exclude.TableStyle.VerticalAlignment==false)
                    {
                        WriteCssItem($"vertical-align:{vAlign};", _settings.Minify);
                    }
                    WriteClassEnd(_settings.Minify);
                }
            }
        }
        internal void AddToCss(string name, ExcelTableStyleElement element, string htmlElement)
        {
            var s = element.Style;
            if (s.HasValue == false) return; //Dont add empty elements
            WriteClass($"table.{name}{htmlElement}{{", _settings.Minify);
            WriteFillStyles(s.Fill);
            WriteFontStyles(s.Font);
            WriteBorderStyles(s.Border);
            WriteClassEnd(_settings.Minify);
        }

        internal void AddHyperlinkCss(string name, ExcelTableStyleElement element)
        {
            WriteClass($"table.{name} a{{", _settings.Minify);
            WriteFontStyles(element.Style.Font);
            WriteClassEnd(_settings.Minify);
        }

        internal void AddToCssBorderVH(string name, ExcelTableStyleElement element, string htmlElement)
        {
            var s = element.Style;
            if (s.Border.Vertical.HasValue == false && s.Border.Horizontal.HasValue==false) return; //Dont add empty elements
            WriteClass($"table.{name}{htmlElement} td,tr {{", _settings.Minify);
            WriteBorderStylesVerticalHorizontal(s.Border);
            WriteClassEnd(_settings.Minify);
        }
        internal void FlushStream()
        {
            _writer.Flush();
        }
        private void WriteFillStyles(ExcelDxfFill f)
        {
            if (f.HasValue && _settings.Css.Exclude.TableStyle.Fill == false)
            {
                if (f.Style == eDxfFillStyle.PatternFill)
                {
                    if (f.PatternType.Value==ExcelFillStyle.Solid)
                    {
                        WriteCssItem($"background-color:{GetDxfColor(f.BackgroundColor)};", _settings.Minify);
                    }
                    else
                    {
                        WriteCssItem($"{PatternFills.GetPatternSvg(f.PatternType.Value, GetDxfColor(f.BackgroundColor), GetDxfColor(f.PatternColor))};", _settings.Minify);
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
            var sb = new StringBuilder();
            if(gradient.GradientType==eDxfGradientFillType.Linear)
            {
                sb.Append($"background: linear-gradient({(gradient.Degree+90)%360}deg");
            }
            else 
            {
                sb.Append($"background:radial-gradient(ellipse {(gradient.Right??0)*100}% {(gradient.Bottom ?? 0) * 100}%");
            }
            foreach (var color in gradient.Colors)
            {
                sb.Append($",{GetDxfColor(color.Color)} {color.Position.ToString("F", CultureInfo.InvariantCulture)}%");
            }
            sb.Append(")");

            WriteCssItem(sb.ToString(), _settings.Minify);
        }
        private void WriteFontStyles(ExcelDxfFontBase f)
        {
            var flags = _settings.Css.Exclude.TableStyle.Font;
            if (f.Color.HasValue && EnumUtil.HasNotFlag(flags, eFontExclude.Color))
            {
                WriteCssItem($"color:{GetDxfColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold.HasValue && f.Bold.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Bold))
            {
                WriteCssItem("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic.HasValue && f.Italic.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Italic))
            {
                WriteCssItem("font-style:italic;", _settings.Minify);
            }
            if (f.Strike.HasValue && f.Strike.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Strike))
            {
                WriteCssItem("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None && EnumUtil.HasNotFlag(flags, eFontExclude.Underline))
            {
                WriteCssItem("text-decoration:underline ", _settings.Minify);
                switch (f.Underline.Value)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        WriteCssItem("double;", _settings.Minify);
                        break;
                    default:
                        WriteCssItem("solid;", _settings.Minify);
                        break;
                }
            }
        }
        private void WriteBorderStyles(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                var flags = _settings.Css.Exclude.TableStyle.Border;
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
                var flags = _settings.Css.Exclude.TableStyle.Border;
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
                var sb = new StringBuilder();
                sb.Append(WriteBorderItemLine(bi.Style.Value, suffix));
                if (bi.Color.HasValue)
                {
                    sb.Append($" {GetDxfColor(bi.Color)}");
                }
                sb.Append(";");

                WriteCssItem(sb.ToString(), _settings.Minify);
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
