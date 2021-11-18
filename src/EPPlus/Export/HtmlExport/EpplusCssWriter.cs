using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System;

namespace OfficeOpenXml.Export.HtmlExport
{
    public class EpplusCssWriter
    {
        readonly Stream _stream;
        readonly StreamWriter _writer;
        readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();
        internal Dictionary<ulong, int> _styleCache;
        const string IndentWhiteSpace = "  ";
        private bool _newLine;
        ExcelRange _range;
        ExcelTheme _theme;
        public EpplusCssWriter(StreamWriter writer, ExcelRangeBase range)
        {
            _stream = writer.BaseStream;
            _writer = writer;

            if (_range.Worksheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                _range.Worksheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = range.Worksheet.Workbook.ThemeManager.CurrentTheme;
        }
        public EpplusCssWriter(Stream stream, ExcelRangeBase range)
        {
            _stream = stream;
            if (_range.Worksheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                _range.Worksheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = range.Worksheet.Workbook.ThemeManager.CurrentTheme;
            _writer = new StreamWriter(stream);
        }
        internal int Indent { get; set; }

        internal void RenderCss(List<string> dataTypes)
        {
            _styleCache = new Dictionary<ulong, int>();
            var styles = _range.Worksheet.Workbook.Styles;            
            var ce = new CellStoreEnumerator<ExcelValue>(_range.Worksheet._values, _range._fromRow, _range._fromCol, _range._toRow, _range._toCol);
            foreach (var c in _range)
            {
                if (c.StyleID > 0 && c.StyleID < styles.CellXfs.Count)
                {
                    AddToCss(styles, c);
                }
            }
        }

        private void AddToCss(ExcelStyles styles, ExcelRangeBase c)
        {
            var xfs = styles.CellXfs[c.StyleID];
            if (xfs.FontId > 0 || xfs.FillId > 0 || xfs.BorderId > 0)
            {
                var key = (ulong)(xfs.FontId << 32 | xfs.BorderId << 16 | xfs.FillId);

                _writer.Write($"style{xfs.Id}");
                _writer.Write("{");
                if (xfs.FillId > 0)
                {
                    WriteFillStyles(xfs.Fill);
                }
                if(xfs.FontId > 0)
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

        private void WriteStyles(ExcelXfs xfs)
        {
            if (xfs.WrapText)
            {
                _writer.Write("word-break: break-word;");
            }
            if(xfs.VerticalAlignment!=ExcelVerticalAlignment.Bottom)
            {

            }
            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General)
            {
                //hAlign = null;
                //_writer.Write($"text-align:{hAlign};");
            }
        }

        private void WriteBorderStyles(ExcelBorderXml b)
        {
            WriteBorderItem(b.Top, "top");
            WriteBorderItem(b.Bottom, "bottom");
            WriteBorderItem(b.Left, "left");
            WriteBorderItem(b.Right, "right");
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");
        }

        private void WriteBorderItem(ExcelBorderItemXml bi, string suffix)
        {
            if (bi.Style != ExcelBorderStyle.None)
            {
                _writer.Write(BorderHelper.WriteBorderItemLine(bi.Style, suffix));
                if (bi.Color!=null && bi.Color.Exists)
                {
                    _writer.Write($" {GetColor(bi.Color)}");
                }
                _writer.Write(";");
            }
        }

        private void WriteFontStyles(ExcelFontXml f)
        {
            if(string.IsNullOrEmpty(f.Name)==false)
            {
                _writer.Write($"font-family:{f.Name};");
            }
            if(f.Size>0)
            {
                _writer.Write($"font-family:{f.Size.ToString("F", CultureInfo.InvariantCulture)};");
            }
            if (f.Color!=null)
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
