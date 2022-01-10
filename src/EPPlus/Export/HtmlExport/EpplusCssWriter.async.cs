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
        internal async Task AddToCssAsync(ExcelStyles styles, int styleId)
        {
            var xfs = styles.CellXfs[styleId];
            if (HasStyle(xfs))
            {
                if (IsAddedToCache(xfs, out int id)==false)
                {
                    await WriteClassAsync($".s{id}{{", _settings.Minify);
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
            if (xfs.WrapText && _cssExclude.WrapText == false)
            {
                await WriteCssItemAsync("word-break: break-word;", _settings.Minify);
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
            if(xfs.TextRotation!=0 && _cssExclude.TextRotation==false)
            {
                await WriteCssItemAsync($"transform: rotate({xfs.TextRotation}deg);", _settings.Minify);
            }

            if(xfs.Indent>0 && _cssExclude.Indent == false)
            {
                await WriteCssItemAsync($"padding-left:{xfs.Indent*_cssSettings.IndentValue}{_cssSettings.IndentUnit};", _settings.Minify);
            }
        }

        private async Task WriteBorderStylesAsync(ExcelBorderXml b)
        {
            await WriteBorderItemAsync(b.Top, "top");
            await WriteBorderItemAsync(b.Bottom, "bottom");
            await WriteBorderItemAsync(b.Left, "left");
            await WriteBorderItemAsync(b.Right, "right");
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
            if(f.Size>0)
            {
                await WriteCssItemAsync($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            if (f.Color!=null && f.Color.Exists)
            {
                await WriteCssItemAsync($"color:{GetColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold)
            {
                await WriteCssItemAsync("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic)
            {
                await WriteCssItemAsync("font-style:italic;", _settings.Minify);
            }
            if (f.Strike)
            {
                await WriteCssItemAsync("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.UnderLineType != ExcelUnderLineType.None)
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
            if(f is ExcelGradientFillXml gf && gf.Type!=ExcelFillGradientType.None)
            {
                await WriteGradientAsync(gf);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
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
