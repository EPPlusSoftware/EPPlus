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
using OfficeOpenXml.Export.HtmlExport.Exporters;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal partial class EpplusTableCssWriter : HtmlWriterBase
    {
        internal async Task RenderAdditionalAndFontCssAsync()
        {
            await WriteClassAsync($"table.{AbstractHtmlExporter.TableClass}{{", _settings.Minify);
            var ns = _table.WorkSheet.Workbook.Styles.GetNormalStyle();
            if (ns != null)
            {
                await WriteCssItemAsync($"font-family:{ns.Style.Font.Name};", _settings.Minify);
                await WriteCssItemAsync($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            foreach (var item in _settings.Css.AdditionalCssElements)
            {
                await WriteCssItemAsync($"{item.Key}:{item.Value};", _settings.Minify);
            }
            await WriteClassEndAsync(_settings.Minify);
        }

        internal async Task AddAlignmentToCssAsync(string name, List<string> dataTypes)
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
                    if(xfs.ApplyAlignment == true)
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
                    await WriteClassAsync($"table.{name} td:nth-child({col}){{", _settings.Minify);
                    if (string.IsNullOrEmpty(hAlign)==false && _settings.Css.Exclude.TableStyle.HorizontalAlignment==false)
                    {
                        await WriteCssItemAsync($"text-align:{hAlign};", _settings.Minify);
                    }
                    if (string.IsNullOrEmpty(vAlign) == false && _settings.Css.Exclude.TableStyle.VerticalAlignment==false)
                    {
                        await WriteCssItemAsync($"vertical-align:{vAlign};", _settings.Minify);
                    }
                    await WriteClassEndAsync(_settings.Minify);
                }
            }
        }
        internal async Task AddToCssAsync(string name, ExcelTableStyleElement element, string htmlElement)
        {
            var s = element.Style;
            if (s.HasValue == false) return; //Dont add empty elements
            await WriteClassAsync($"table.{name}{htmlElement}{{", _settings.Minify);
            await WriteFillStylesAsync(s.Fill);
            await WriteFontStylesAsync(s.Font);
            await WriteBorderStylesAsync(s.Border);
            await WriteClassEndAsync(_settings.Minify);
        }

        internal async Task AddHyperlinkCssAsync(string name, ExcelTableStyleElement element)
        {
            await WriteClassAsync($"table.{name} a{{", _settings.Minify);
            await WriteFontStylesAsync(element.Style.Font);
            await WriteClassEndAsync(_settings.Minify);
        }

        internal async Task AddToCssBorderVHAsync(string name, ExcelTableStyleElement element, string htmlElement)
        {
            var s = element.Style;
            if (s.Border.Vertical.HasValue == false && s.Border.Horizontal.HasValue==false) return; //Dont add empty elements
            await WriteClassAsync($"table.{name}{htmlElement} td,tr {{", _settings.Minify);
            await WriteBorderStylesVerticalHorizontalAsync(s.Border);
            await WriteClassEndAsync(_settings.Minify);
        }
        internal async Task FlushStreamAsync()
        {
            await _writer.FlushAsync();
        }
        private async Task WriteFillStylesAsync(ExcelDxfFill f)
        {
            if (f.HasValue && _settings.Css.Exclude.TableStyle.Fill == false)
            {
                if (f.Style == eDxfFillStyle.PatternFill)
                {
                    if (f.PatternType.Value==ExcelFillStyle.Solid)
                    {
                        await WriteCssItemAsync($"background-color:{GetDxfColor(f.BackgroundColor)};", _settings.Minify);
                    }
                    else
                    {
                        await WriteCssItemAsync($"{PatternFills.GetPatternSvg(f.PatternType.Value, GetDxfColor(f.BackgroundColor), GetDxfColor(f.PatternColor))};", _settings.Minify);
                    }
                }
                else if(f.Style==eDxfFillStyle.GradientFill)
                {
                    await WriteDxfGradientAsync(f.Gradient);
                }
            }
        }

        private async Task WriteDxfGradientAsync(ExcelDxfGradientFill gradient)
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

            await WriteCssItemAsync(sb.ToString(), _settings.Minify);
        }
        private async Task WriteFontStylesAsync(ExcelDxfFontBase f)
        {
            var flags = _settings.Css.Exclude.TableStyle.Font;
            if (f.Color.HasValue && EnumUtil.HasNotFlag(flags, eFontExclude.Color))
            {
                await WriteCssItemAsync($"color:{GetDxfColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold.HasValue && f.Bold.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Bold))
            {
                await WriteCssItemAsync("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic.HasValue && f.Italic.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Italic))
            {
                await WriteCssItemAsync("font-style:italic;", _settings.Minify);
            }
            if (f.Strike.HasValue && f.Strike.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Strike))
            {
                await WriteCssItemAsync("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None && EnumUtil.HasNotFlag(flags, eFontExclude.Underline))
            {
                await WriteCssItemAsync("text-decoration:underline ", _settings.Minify);
                switch (f.Underline.Value)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        await WriteCssItemAsync("double;", _settings.Minify);
                        break;
                    default:
                        await WriteCssItemAsync("solid;", _settings.Minify);
                        break;
                }
            }
        }
        private async Task WriteBorderStylesAsync(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                var flags = _settings.Css.Exclude.TableStyle.Border;
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Top)) await WriteBorderItemAsync(b.Top, "top");
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom)) await WriteBorderItemAsync(b.Bottom, "bottom");
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Left)) await WriteBorderItemAsync(b.Left, "left");
                if(EnumUtil.HasNotFlag(flags, eBorderExclude.Right)) await WriteBorderItemAsync(b.Right, "right");
            }
        }
        private async Task WriteBorderStylesVerticalHorizontalAsync(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                var flags = _settings.Css.Exclude.TableStyle.Border;
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Top)) await WriteBorderItemAsync(b.Horizontal, "top");
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom)) await WriteBorderItemAsync(b.Horizontal, "bottom");
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Left)) await WriteBorderItemAsync(b.Vertical, "left");
                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Right)) await WriteBorderItemAsync(b.Vertical, "right");
            }
        }

        private async Task WriteBorderItemAsync(ExcelDxfBorderItem bi, string suffix)
        {
            if (bi.HasValue && bi.Style != ExcelBorderStyle.None)
            {
                var sb = new StringBuilder();
                sb.Append(GetBorderItemLine(bi.Style.Value, suffix));
                if (bi.Color.HasValue)
                {
                    sb.Append($" {GetDxfColor(bi.Color)}");
                }
                sb.Append(";");

                await WriteCssItemAsync(sb.ToString(), _settings.Minify);
            }
        }
    }
#endif
}
