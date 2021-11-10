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
namespace OfficeOpenXml.Export.HtmlExport
{
    internal class EpplusCssWriter
    {
        readonly Stream _stream;
        readonly StreamWriter _writer;
        readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        const string IndentWhiteSpace = "  ";
        private bool _newLine;
        ExcelTable _table;
        ExcelTheme _theme;
        public EpplusCssWriter(Stream stream, ExcelTable table)
        {
            _stream = stream;
            _table = table;
            if(table.WorkSheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                table.WorkSheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = table.WorkSheet.Workbook.ThemeManager.CurrentTheme;
            _writer = new StreamWriter(stream);
        }
        internal int Indent { get; set; }

        internal void RenderCss() 
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

            var tableClass = $"epplus-tablestyle-{tblStyle.Name.ToLower()}";
            AddToCss($"{tableClass}", tblStyle.WholeTable, "");

            if(_table.ShowHeader)
            {
                AddToCss($"{tableClass}", tblStyle.HeaderRow, " thead tr th");
                AddToCss($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:nth-child(1)");
                if (_table.Columns.Count > 1)
                {
                    AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:nth-child({_table.Columns.Count})");
                }
            }

            if (_table.ShowTotal)
            {
                AddToCss($"{tableClass}", tblStyle.TotalRow, " tfoot tr td");
                AddToCss($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:nth-child(1)");
                if (_table.Columns.Count > 1)
                {
                    AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:nth-child({_table.Columns.Count})");
                }
            }

            if (_table.ShowFirstColumn)
            {
                AddToCss($"{tableClass}", tblStyle.FirstColumn, " tbody tr td:nth-child(1)");                
            }

            if (_table.ShowLastColumn && _table.Columns.Count > 1)
            {
                AddToCss($"{tableClass}", tblStyle.FirstColumn, $" tbody tr td:nth-child({_table.Columns.Count})");
            }

            if(_table.ShowColumnStripes)
            {
                AddToCss($"{tableClass}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
                AddToCss($"{tableClass}", tblStyle.SecondColumnStripe.Style.HasValue ? tblStyle.SecondColumnStripe : tblStyle.WholeTable, $" tbody tr td:nth-child(even)");
            }

            if (_table.ShowRowStripes)
            {
                AddToCss($"{tableClass}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd) td");
                AddToCss($"{tableClass}", tblStyle.SecondRowStripe.Style.HasValue ? tblStyle.SecondRowStripe : tblStyle.WholeTable, " tbody tr:nth-child(even) td");
            }
            else
            {
                AddToCss($"{tableClass}", tblStyle.FirstRowStripe, " thead tr td");
            }

            _writer.Flush();
        }
        private void AddToCss(string name, ExcelTableStyleElement element, string htmlElement)
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

        private void WriteFillStyles(ExcelDxfFill f)
        {
            if (f.HasValue)
            {
                if (f.Style == eDxfFillStyle.PatternFill)
                {
                    _writer.Write($"background-color:{GetDxfColor(f.PatternColor)};");
                }
            }
        }

        private void WriteFontStyles(ExcelDxfFontBase f)
        {
            if (f.Color.HasValue)
            {
                _writer.Write($"color:{GetDxfColor(f.Color)};");
                //color: #007731;
            }
            if (f.Bold.HasValue && f.Bold.Value)
            {
                _writer.Write("font-weight:bolder;");
            }
            if (f.Italic.HasValue && f.Italic.Value)
            {
                _writer.Write("font-style:italic;");
            }
            if (f.Strike.HasValue && f.Strike.Value)
            {
                _writer.Write("text-decoration:line-through solid;");
            }
            if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None)
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
            if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None)
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
                WriteBorderItem(b.Top, "top");
                WriteBorderItem(b.Bottom, "bottom");
                WriteBorderItem(b.Left, "left");
                WriteBorderItem(b.Right, "right");
            }
        }

        private void WriteBorderItem(ExcelDxfBorderItem bi, string suffix)
        {
            if (bi.HasValue && bi.Style != ExcelBorderStyle.None)
            {
                _writer.Write($"border-{suffix}:");
                switch (bi.Style)
                {
                    case ExcelBorderStyle.Hair:
                        _writer.Write($"1px solid");
                        break;
                    case ExcelBorderStyle.Thin:
                        _writer.Write($"thin solid");
                        break;
                    case ExcelBorderStyle.Medium:
                        _writer.Write($"medium solid");
                        break;
                    case ExcelBorderStyle.Thick:
                        _writer.Write($"thick solid");
                        break;
                    case ExcelBorderStyle.Double:
                        _writer.Write($"double");
                        break;
                    case ExcelBorderStyle.Dotted:
                        _writer.Write($"dotted");
                        break;
                    case ExcelBorderStyle.Dashed:
                    case ExcelBorderStyle.DashDot:
                    case ExcelBorderStyle.DashDotDot:
                        _writer.Write($"dashed");
                        break;
                    case ExcelBorderStyle.MediumDashed:
                    case ExcelBorderStyle.MediumDashDot:
                    case ExcelBorderStyle.MediumDashDotDot:
                        _writer.Write($"medium dashed");
                        break;
                }
                if (bi.Color.HasValue)
                {
                    _writer.Write($" {GetDxfColor(bi.Color)}");
                }
                _writer.Write(";");
            }
        }
        private string GetDxfColor(Style.Dxf.ExcelDxfColor c)
        {
            Color ret;
            if (c.Color.HasValue)
            {
                ret = c.Color.Value;
            }
            else if (c.Theme.HasValue)
            {
                ret = GetThemeColor(c.Theme.Value);
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
                ret = ApplyTint(ret, c.Tint.Value);
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }

        internal Color ApplyTint(Color ret, double tint)
        {
            if (tint == 0)
            {
                return ret;
            }
            else
            {
                ExcelDrawingRgbColor.GetHslColor(ret, out double h, out double s, out double l);
                if (tint < 0)
                {
                    l = l*(1.0 + tint);
                }
                else if (tint > 0)
                {
                    //l = (1-l)*tint;
                    l = 1 - l * (1 - tint);
                }
                return ExcelDrawingHslColor.GetRgb(h, s, l);
            }
        }

        private Color GetThemeColor(eThemeSchemeColor tc)
        {
            var cm = _theme.ColorScheme.GetColorByEnum(tc);
            return GetThemeColor(cm);
        }

        private Color GetThemeColor(ExcelDrawingThemeColorManager cm)
        {
            Color color;
            switch (cm.ColorType)
            {
                case eDrawingColorType.Rgb:
                    color = cm.RgbColor.Color;
                    break;
                case eDrawingColorType.Preset:
                    color = Color.FromName(cm.PresetColor.Color.ToString());
                    break;
                case eDrawingColorType.System:
                    color = cm.SystemColor.GetColor();
                    break;
                case eDrawingColorType.RgbPercentage:
                    var rp = cm.RgbPercentageColor;
                    color = Color.FromArgb(GetRgpPercentToRgb(rp.RedPercentage),
                                   GetRgpPercentToRgb(rp.GreenPercentage),
                                   GetRgpPercentToRgb(rp.BluePercentage));
                    break;
                case eDrawingColorType.Hsl:
                    color = cm.HslColor.GetRgbColor();
                    break;
                default:
                    color = Color.Empty;
                    break;
            }

            //TODO:Apply Transforms

            return color;
        }

        private int GetRgpPercentToRgb(double percentage)
        {
            if (percentage < 0) return 0;
            if (percentage > 255) return 255;
            return (int)(percentage * 255 / 100);
        }

    }
}
