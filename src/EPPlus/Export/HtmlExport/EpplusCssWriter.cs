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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System;
using OfficeOpenXml.Utils;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class EpplusCssWriter : HtmlWriterBase
    {
        protected HtmlTableExportSettings _settings;
        ExcelRangeBase _range;
        ExcelTheme _theme;
        internal eFontExclude _fontExclude;
        internal eBorderExclude _borderExclude;
        internal EpplusCssWriter(StreamWriter writer, ExcelRangeBase range, HtmlTableExportSettings settings) : base(writer) 
        {
            _settings = settings;
            Init(range);
        }
        internal EpplusCssWriter(Stream stream, ExcelRangeBase range, HtmlTableExportSettings settings) : base(stream, settings.Encoding)
        {
            _settings = settings;
            Init(range);
        }
        private void Init(ExcelRangeBase range)
        {
            _range = range;

            if (_range.Worksheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                _range.Worksheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = range.Worksheet.Workbook.ThemeManager.CurrentTheme;
            _borderExclude = _settings.Css.Exclude.CellStyle.Border;
            _fontExclude = _settings.Css.Exclude.CellStyle.Font;
        }

        internal void AddToCss(ExcelStyles styles, int styleId)
        {
            var xfs = styles.CellXfs[styleId];
            if (HasStyle(xfs))
            {
                if (IsAddedToCache(xfs, out int id)==false)
                {
                    WriteClass($".s{id}{{", _settings.Minify);
                    if (xfs.FillId > 0)
                    {
                        WriteFillStyles(xfs.Fill);
                    }
                    if (xfs.FontId > 0)
                    {
                        WriteFontStyles(xfs.Font);
                    }
                    if (xfs.BorderId > 0)
                    {
                        WriteBorderStyles(xfs.Border);
                    }
                    WriteStyles(xfs);
                    WriteClassEnd(_settings.Minify);
                }
            }
        }

        private bool IsAddedToCache(ExcelXfs xfs, out int id)
        {
            var key = GetStyleKey(xfs);
            if (_styleCache.ContainsKey(key))
            {
                id = _styleCache[key];
                return true;
            }
            else
            {
                id = _styleCache.Count+1;
                _styleCache.Add(key, id);
                return false;
            }
        }

        private void WriteStyles(ExcelXfs xfs)
        {
            if (xfs.WrapText && _settings.Css.Exclude.CellStyle.WrapText == false)
            {
                WriteCssItem("word-break: break-word;", _settings.Minify);
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && _settings.Css.Exclude.CellStyle.HorizontalAlignment == false)
            {
                var hAlign = GetHorizontalAlignment(xfs);
                WriteCssItem($"text-align:{hAlign};", _settings.Minify);
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && _settings.Css.Exclude.CellStyle.VerticalAlignment == false)
            {
                var vAlign = GetVerticalAlignment(xfs);
                WriteCssItem($"vertical-align:{vAlign};", _settings.Minify);
            }
            if(xfs.TextRotation!=0 && _settings.Css.Exclude.CellStyle.TextRotation==false)
            {
                WriteCssItem($"transform: rotate({xfs.TextRotation}deg);", _settings.Minify);
            }

            if(xfs.Indent>0 && _settings.Css.Exclude.CellStyle.Indent == false)
            {
                WriteCssItem($"padding-left:{xfs.Indent*_settings.Css.IndentValue}{_settings.Css.IndentUnit};", _settings.Minify);
            }
        }

        private void WriteBorderStyles(ExcelBorderXml b)
        {
            WriteBorderItem(b.Top, "top");
            WriteBorderItem(b.Bottom, "bottom");
            WriteBorderItem(b.Left, "left");
            WriteBorderItem(b.Right, "right");
            //TODO add Diagonal
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");
        }

        private void WriteBorderItem(ExcelBorderItemXml bi, string suffix)
        {
            if (bi.Style != ExcelBorderStyle.None)
            {
                var sb = new StringBuilder();
                sb.Append(WriteBorderItemLine(bi.Style, suffix));
                if (bi.Color!=null && bi.Color.Exists)
                {
                    sb.Append($" {GetColor(bi.Color)}");
                }
                sb.Append(";");

                WriteCssItem(sb.ToString(), _settings.Minify);
            }
        }

        private void WriteFontStyles(ExcelFontXml f)
        {
            if(string.IsNullOrEmpty(f.Name)==false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name))
            {
                WriteCssItem($"font-family:{f.Name};", _settings.Minify);
            }
            if(f.Size>0)
            {
                WriteCssItem($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            if (f.Color!=null && f.Color.Exists)
            {
                WriteCssItem($"color:{GetColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold)
            {
                WriteCssItem("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic)
            {
                WriteCssItem("font-style:italic;", _settings.Minify);
            }
            if (f.Strike)
            {
                WriteCssItem("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.UnderLineType != ExcelUnderLineType.None)
            {
                switch (f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        WriteCssItem("text-decoration:underline double;", _settings.Minify);
                        break;
                    default:
                        WriteCssItem("text-decoration:underline solid;", _settings.Minify);
                        break;
                }
            }            
        }

        private void WriteFillStyles(ExcelFillXml f)
        {
            if(f is ExcelGradientFillXml gf && gf.Type!=ExcelFillGradientType.None)
            {
                WriteGradient(gf);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
                {
                    if (f.PatternType == ExcelFillStyle.Solid)
                    {
                        WriteCssItem($"background-color:{GetColor(f.BackgroundColor)};", _settings.Minify);
                    }
                    else
                    {
                        WriteCssItem($"{PatternFills.GetPatternSvg(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor))}", _settings.Minify);
                    }
                }
            }
        }

        private void WriteGradient(ExcelGradientFillXml gradient)
        {
            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                _writer.Write($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                _writer.Write($"background:radial-gradient(ellipse {gradient.Right  * 100}% {gradient.Bottom  * 100}%");
            }

            _writer.Write($",{GetColor(gradient.GradientColor1)} 0%");
            _writer.Write($",{GetColor(gradient.GradientColor2)} 100%");

            _writer.Write(");");
        }
        private string GetColor(ExcelColorXml c)
        {
            Color ret;
            if (!string.IsNullOrEmpty(c.Rgb))
            {
                if (int.TryParse(c.Rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (c.Theme.HasValue)
            {
                ret = ColorConverter.GetThemeColor(_theme, c.Theme.Value);
            }
            else if (c.Indexed >= 0)
            {
                ret = ExcelColor.GetIndexedColor(c.Indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (c.Tint != 0)
            {
                ret = ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
        public void FlushStream()
        {
            _writer.Flush();
        }
    }
}
