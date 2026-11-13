using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tc = OfficeOpenXml.Utils.TypeConversion;
namespace OfficeOpenXml.Drawing.Renderer.Chart.Defaults
{
    internal static class DefaultChartBorder
    {
        internal static (Color color, double width) GetDefaulDataPoint2d(ExcelTheme theme, ExcelChart chart, int serieIndex)
        {
            (Color color, double width) ret = new();
            var chartStyle = (int)chart.Style;
            if (chartStyle >= 9 && chartStyle <= 16)
            {
                ret.width = chart.IsTypeLine() ? 5 : theme.FormatScheme.BorderStyle[0].Width;
                ret.color = theme.ColorScheme.Light1.GetColor();
            }
            else if (chartStyle >= 33 && chartStyle <= 40)
            {
                ret.width = chart.IsTypeLine() ? 5 : (double)theme.FormatScheme.BorderStyle[0].Width / ExcelDrawing.EMU_PER_POINT;
                switch (chartStyle)
                {
                    case 33: //50 % shade of dk
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Dark1.GetColor(), -0.5);
                        break;
                    case 34:  //Pattern 3
                        ret.color = GetColorPattern3(theme, serieIndex);
                        break;
                    case 35:  //50 % shade of accent1
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Accent1.GetColor(), -0.5);
                        break;
                    case 36:  //50 % shade of accent2
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Accent2.GetColor(), -0.5);
                        break;
                    case 37:  //50 % shade of accent3
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Accent3.GetColor(), -0.5);
                        break;
                    case 38:  //50 % shade of accent4
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Accent4.GetColor(), -0.5);
                        break;
                    case 39:  //50 % shade of accent5
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Accent5.GetColor(), -0.5);
                        break;
                    case 40:  //50 % shade of accent6
                        ret.color = tc.ColorConverter.ApplyTintDrawing(theme.ColorScheme.Accent6.GetColor(), -0.5);
                        break;
                }
            }
            else
            {
                ret.width = 0;
                ret.color = Color.Empty;
            }
            return ret;
        }

        private static Color GetColorPattern3(ExcelTheme theme, int serieIndex)
        {
            Color bc = theme.ColorScheme.GetColorByEnum(eSchemeColor.Accent1+(serieIndex%6)).GetColor();
            var variations = ExcelColorTransformCollection.GetDefault();
            var ix = ((serieIndex - (serieIndex % 6)) / 6) % (variations.Count+1);
            var color = tc.ColorConverter.ApplyTintDrawing(bc, -0.5);
            if(ix>0)
            {
                return tc.ColorConverter.ApplyTransforms(color, variations[ix-1]);
            }
            return color;
        }
    }
}
