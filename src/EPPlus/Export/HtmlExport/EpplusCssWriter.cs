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

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class EpplusCssWriter : CssWriterBase
    {
        readonly Stream _stream;
        readonly StreamWriter _writer;
        readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();
        internal Dictionary<ulong, int> _styleCache = new Dictionary<ulong, int>();
        const string IndentWhiteSpace = "  ";
        private bool _newLine;
        ExcelRangeBase _range;
        ExcelTheme _theme;
        internal eFontExclude _fontExclude;
        internal eBorderExclude _borderExclude;
        internal EpplusCssWriter(StreamWriter writer, ExcelRangeBase range, CssTableExportOptions options)
        {
            _stream = writer.BaseStream;
            _writer = writer;
            _options = options;
            Init(range);
        }
        internal EpplusCssWriter(Stream stream, ExcelRangeBase range, CssTableExportOptions options)
        {
            _stream = stream;
            _writer = new StreamWriter(stream);
            _options = options;
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
        }

        internal void RenderCss(List<string> dataTypes)
        {
            var styles = _range.Worksheet.Workbook.Styles;
            _borderExclude = _options.Exclude.CellStyle.Border;
            _fontExclude = _options.Exclude.CellStyle.Font; 
            var ce = new CellStoreEnumerator<ExcelValue>(_range.Worksheet._values, _range._fromRow, _range._fromCol, _range._toRow, _range._toCol);
            while(ce.Next())
            {
                if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                {
                    AddToCss(styles, ce.Value._styleId);
                }
            }
            _writer.Flush();
        }

        private void AddToCss(ExcelStyles styles, int styleId)
        {
            var xfs = styles.CellXfs[styleId];
            if (xfs.FontId > 0 || xfs.FillId > 0 || xfs.BorderId > 0)
            {
                int id = GetOrAddToStyleCache(styleId, xfs);
                _writer.Write($".s{id}");
                _writer.Write("{");
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
                _writer.Write("}");
            }
        }

        private int GetOrAddToStyleCache(int styleId, ExcelXfs xfs)
        {
            var key = (ulong)(xfs.FontId << 32 | xfs.BorderId << 16 | xfs.FillId);
            int id;
            if (_styleCache.ContainsKey(key))
            {
                id = _styleCache[key];
            }
            else
            {
                id = _styleCache.Count+1;
                _styleCache.Add(key, id);
            }

            return id;
        }

        private void WriteStyles(ExcelXfs xfs)
        {
            if (xfs.WrapText && _options.Exclude.CellStyle.WrapText == false)
            {
                _writer.Write("word-break: break-word;");
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && _options.Exclude.CellStyle.HorizontalAlignment == false)
            {
                var hAlign = GetHorizontalAlignment(xfs);
                _writer.Write($"text-align:{hAlign};");
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && _options.Exclude.CellStyle.VerticalAlignment == false)
            {
                var vAlign = GetVerticalAlignment(xfs);
                _writer.Write($"vertical-align:{vAlign};");
            }
            if(xfs.TextRotation!=0 && _options.Exclude.CellStyle.TextRotation==false)
            {
                _writer.Write($"transform: rotate({xfs.TextRotation}deg);");
            }

            if(xfs.Indent>0 && _options.Exclude.CellStyle.Indent == false)
            {
                _writer.Write($"padding-left:{xfs.Indent*_options.Indent}{_options.IndentUnit};"); //
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
                _writer.Write(WriteBorderItemLine(bi.Style, suffix));
                if (bi.Color!=null && bi.Color.Exists)
                {
                    _writer.Write($" {GetColor(bi.Color)}");
                }
                _writer.Write(";");
            }
        }

        private void WriteFontStyles(ExcelFontXml f)
        {
            if(string.IsNullOrEmpty(f.Name)==false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name))
            {
                _writer.Write($"font-family:{f.Name};");
            }
            if(f.Size>0)
            {
                _writer.Write($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;");
            }
            if (f.Color!=null && f.Color.Exists)
            {
                _writer.Write($"color:{GetColor(f.Color)};");
            }
            if (f.Bold)
            {
                _writer.Write("font-weight:bolder;");
            }
            if (f.Italic)
            {
                _writer.Write("font-style:italic;");
            }
            if (f.Strike)
            {
                _writer.Write("text-decoration:line-through solid;");
            }
            if (f.UnderLineType != ExcelUnderLineType.None)
            {
                _writer.Write("text-decoration:underline ");
                switch (f.UnderLineType)
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
                        _writer.Write($"background-color:{GetColor(f.BackgroundColor)};");
                    }
                    else
                    {
                        _writer.Write($"{PatternFills.GetPatternSvg(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor))}");
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
    }
}
