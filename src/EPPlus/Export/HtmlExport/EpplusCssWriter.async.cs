/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/17/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System;
using OfficeOpenXml.Utils;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal partial class EpplusCssWriter : HtmlWriterBase
    {
        internal async Task RenderAdditionalAndFontCssAsync(string tableClass)
        {
            await WriteClassAsync($"table.{tableClass}{{", _settings.Minify);
            var ns = _range.Worksheet.Workbook.Styles.GetNormalStyle();
            if (ns != null)
            {
                await WriteCssItemAsync($"font-family:{ns.Style.Font.Name};", _settings.Minify);
                await WriteCssItemAsync($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            foreach (var item in _cssSettings.AdditionalCssElements)
            {
                await WriteCssItemAsync($"{item.Key}:{item.Value};", _settings.Minify);
            }
            await WriteClassEndAsync(_settings.Minify);
            if (_settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType ||
                _settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.CellDataType)
            {
                await WriteClassAsync($".{_settings.StyleClassPrefix}al {{", _settings.Minify);
                await WriteCssItemAsync($"text-align:left;", _settings.Minify);
                await WriteClassEndAsync(_settings.Minify);
                await WriteClassAsync($".{_settings.StyleClassPrefix}ar {{", _settings.Minify);
                await WriteCssItemAsync($"text-align:right;", _settings.Minify);
                await WriteClassEndAsync(_settings.Minify);
            }
            if (_settings.SetColumnWidth)
            {
                var ws = _range.Worksheet;
                await WriteClassAsync($".{_settings.StyleClassPrefix}dcw {{", _settings.Minify);
                await WriteCssItemAsync($"width:{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px;", _settings.Minify);
                await WriteClassEndAsync(_settings.Minify);

                await WriteClassAsync($".{_settings.StyleClassPrefix}drh {{", _settings.Minify);
                await WriteCssItemAsync($"height:{(int)(ws.DefaultRowHeight / 0.75)}px;", _settings.Minify);
                await WriteClassEndAsync(_settings.Minify);
            }
        }

        internal async Task AddToCssAsync(ExcelStyles styles, int styleId, string styleClassPrefix)
        {
            var xfs = styles.CellXfs[styleId];
            if (HasStyle(xfs))
            {
                if (IsAddedToCache(xfs, out int id)==false)
                {
                    await WriteClassAsync($".{styleClassPrefix}s{id}{{", _settings.Minify);
                    if (xfs.FillId > 0)
                    {
                        await WriteFillStylesAsync(xfs.Fill);
                    }
                    if (xfs.FontId > 0)
                    {
                        await WriteFontStylesAsync(xfs.Font);
                    }
                    if (xfs.BorderId > 0)
                    {
                        await WriteBorderStylesAsync(xfs.Border);
                    }
                    await WriteStylesAsync(xfs);
                    await WriteClassEndAsync(_settings.Minify);
                }
            }
        }
        private async Task WriteStylesAsync (ExcelXfs xfs)
        {
            if (_cssExclude.WrapText == false)
            {
                if (xfs.WrapText)
                {
                    await WriteCssItemAsync("white-space: break-spaces;", _settings.Minify);
                }
                else
                {
                    await WriteCssItemAsync("white-space: nowrap;", _settings.Minify);
                }
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && _cssExclude.HorizontalAlignment == false)
            {
                var hAlign = GetHorizontalAlignment(xfs);
                await WriteCssItemAsync($"text-align:{hAlign};", _settings.Minify);
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && _cssExclude.VerticalAlignment == false)
            {
                var vAlign = GetVerticalAlignment(xfs);
                await WriteCssItemAsync($"vertical-align:{vAlign};", _settings.Minify);
            }
            if (xfs.TextRotation != 0 && _cssExclude.TextRotation == false)
            {
                await WriteCssItemAsync($"transform: rotate({xfs.TextRotation}deg);", _settings.Minify);
            }

            if (xfs.Indent > 0 && _cssExclude.Indent == false)
            {
                await WriteCssItemAsync($"padding-left:{xfs.Indent * _cssSettings.IndentValue}{_cssSettings.IndentUnit};", _settings.Minify);
            }
        }

        private async Task WriteBorderStylesAsync(ExcelBorderXml b)
        {
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Top)) await WriteBorderItemAsync(b.Top, "top");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Bottom)) await WriteBorderItemAsync(b.Bottom, "bottom");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Left)) await WriteBorderItemAsync(b.Left, "left");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Right)) await WriteBorderItemAsync(b.Right, "right");
            //TODO add Diagonal
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");
        }

        private async Task WriteBorderItemAsync(ExcelBorderItemXml bi, string suffix)
        {
            if (bi.Style != ExcelBorderStyle.None)
            {
                var sb = new StringBuilder();
                sb.Append(GetBorderItemLine(bi.Style, suffix));
                if (bi.Color!=null && bi.Color.Exists)
                {
                    sb.Append($" {GetColor(bi.Color)}");
                }
                sb.Append(";");

                await WriteCssItemAsync(sb.ToString(), _settings.Minify);
            }
        }

        private async Task WriteFontStylesAsync(ExcelFontXml f)
        {
            if(string.IsNullOrEmpty(f.Name)==false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name))
            {
                await WriteCssItemAsync($"font-family:{f.Name};", _settings.Minify);
            }
            if(f.Size>0 && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Size))
            {
                await WriteCssItemAsync($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            if (f.Color!=null && f.Color.Exists && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color))
            {
                await WriteCssItemAsync($"color:{GetColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold))
            {
                await WriteCssItemAsync("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic))
            {
                await WriteCssItemAsync("font-style:italic;", _settings.Minify);
            }
            if (f.Strike && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike))
            {
                await WriteCssItemAsync("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline))
            {
                switch (f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        await WriteCssItemAsync("text-decoration:underline double;", _settings.Minify);
                        break;
                    default:
                        await WriteCssItemAsync("text-decoration:underline solid;", _settings.Minify);
                        break;
                }
            }            
        }

        private async Task WriteFillStylesAsync(ExcelFillXml f)
        {
            if (_cssExclude.Fill) return;
            if (f is ExcelGradientFillXml gf && gf.Type!=ExcelFillGradientType.None)
            {
                await WriteGradientAsync(gf);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
                {
                    await WriteCssItemAsync($"background-color:{GetColor(f.BackgroundColor)};", _settings.Minify);
                }
                else
                {
                    await WriteCssItemAsync($"{PatternFills.GetPatternSvg(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor))}", _settings.Minify);
                }
            }
        }

        private async Task WriteGradientAsync(ExcelGradientFillXml gradient)
        {
            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                await _writer.WriteAsync($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                await _writer.WriteAsync($"background:radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%");
            }

            await _writer.WriteAsync($",{GetColor(gradient.GradientColor1)} 0%");
            await _writer.WriteAsync($",{GetColor(gradient.GradientColor2)} 100%");

            await _writer.WriteAsync(");");
        }
        public async Task FlushStreamAsync()
        {
            await _writer.FlushAsync();
        }
    }
#endif
}
