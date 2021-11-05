using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
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
            if(c.Color.HasValue)
            {
                return "#" + c.Color.Value.ToArgb().ToString("x");
            }
            else if(c.Theme.HasValue)
            {
                return GetThemeColor(c.Theme.Value);
            }
            if (c.Tint.HasValue)
            {
                
            }
            return null;
        }

        private string GetThemeColor(eThemeSchemeColor tc)
        {
            var t = _table.WorkSheet.Workbook.ThemeManager.CurrentTheme;
            var cm = t.ColorScheme.GetColorByEnum(tc);
            return GetThemeColor(cm);
        }

        private string GetThemeColor(ExcelDrawingThemeColorManager cm)
        {
            switch(cm.ColorType)
            {
                case eDrawingColorType.Rgb:
                    return "#" + cm.RgbColor.Color.ToArgb().ToString("x");
                case eDrawingColorType.Preset:
                    return cm.PresetColor.Color.ToString();
                case eDrawingColorType.System:
                    var c = cm.SystemColor.GetColor();
                    return "#" + cm.RgbColor.Color.ToArgb().ToString("x");
                case eDrawingColorType.RgbPercentage:
                    var rp = cm.RgbPercentageColor;
                    return "#" + System.Drawing.Color.FromArgb(GetRgpPercentToRgb(rp.RedPercentage),
                                                         GetRgpPercentToRgb(rp.GreenPercentage),
                                                         GetRgpPercentToRgb(rp.BluePercentage)).ToArgb().ToString("x");
                case eDrawingColorType.Hsl:
                    return "#" + cm.HslColor.GetRgbColor().ToArgb().ToString("x");
                    
                
                default:
                    return null;

                    
            }
        }

        private int GetRgpPercentToRgb(double percentage)
        {
            if (percentage < 0) return 0;
            if (percentage > 255) return 255;
            return (int)(percentage * 255 / 100);
        }
    }
}
