using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class TableStyleToCss
    {
        ExcelTable _table;
        internal TableStyleToCss(ExcelTable table)
        {
            _table = table;
        }
        internal void Render(StreamWriter sr)
        {
            if(_table.TableStyle==TableStyles.None)
            {
                return;
            }
            ExcelTableNamedStyleBase tblStyle;
            if(_table.TableStyle==TableStyles.Custom)
            {
                tblStyle = _table.WorkSheet.Workbook.Styles.TableStyles[_table.StyleName];
            }
            else
            {
                var tmpNode = _table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tblStyle = new ExcelTableNamedStyle(_table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, _table.WorkSheet.Workbook.Styles);
                tblStyle.SetFromTemplate(_table.TableStyle);
            }

            AddToCss($"{tblStyle.Name}-HeaderRow", tblStyle.HeaderRow);
        }

        private void AddToCss(string name, ExcelTableStyleElement element)
        {
            var s = element.Style;
            if (s.Fill.Style == eDxfFillStyle.PatternFill)
            {
                var fillColor = GetDxfColor(s.Fill.PatternColor);
            }
        }

        private string GetDxfColor(Style.Dxf.ExcelDxfColor c)
        {
            Color ret;
            if(c.Color.HasValue)
            {
                ret = c.Color.Value;
            }
            else if(c.Theme.HasValue)
            {
                ret = GetThemeColor(c.Theme.Value);
            }
            else if (c.Index.HasValue)
            {
                ret=ExcelColor.GetIndexedColor(c.Index.Value);
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
            return "#" + ret.ToArgb().ToString("x");
        }

        private Color ApplyTint(Color ret, double? tint)
        {
            return ret;
        }

        private Color GetThemeColor(eThemeSchemeColor tc)
        {
            var t = _table.WorkSheet.Workbook.ThemeManager.CurrentTheme;
            var cm = t.ColorScheme.GetColorByEnum(tc);
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
