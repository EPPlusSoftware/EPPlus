/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfThemeColorMath
    {
        internal static System.Drawing.Color ApplyThemeTint(System.Drawing.Color c, double tint)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double max = System.Math.Max(r, System.Math.Max(g, b));
            double min = System.Math.Min(r, System.Math.Min(g, b));
            double h = 0, s = 0, l = (max + min) / 2.0;
            if (max != min)
            {
                double d = max - min;
                s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
                if (max == r) h = (g - b) / d + (g < b ? 6.0 : 0.0);
                else if (max == g) h = (b - r) / d + 2.0;
                else h = (r - g) / d + 4.0;
                h /= 6.0;
            }
            l = tint < 0 ? l * (1.0 + tint) : l * (1.0 - tint) + tint;
            double R, G, B;
            if (s == 0) { R = G = B = l; }
            else
            {
                double q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
                double p = 2.0 * l - q;
                R = Hue2Rgb(p, q, h + 1.0 / 3.0);
                G = Hue2Rgb(p, q, h);
                B = Hue2Rgb(p, q, h - 1.0 / 3.0);
            }
            return System.Drawing.Color.FromArgb(c.A, Clamp255(R), Clamp255(G), Clamp255(B));
        }

        private static double Hue2Rgb(double p, double q, double t)
        {
            if (t < 0) t += 1.0;
            if (t > 1) t -= 1.0;
            if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
            if (t < 1.0 / 2.0) return q;
            if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
            return p;
        }

        private static int Clamp255(double v)
        {
            int x = (int)System.Math.Round(v * 255.0);
            return x < 0 ? 0 : (x > 255 ? 255 : x);
        }
    }
}
