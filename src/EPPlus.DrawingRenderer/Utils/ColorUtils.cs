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
using EPPlus.DrawingRenderer;
using System.Drawing;

namespace EPPlusImageRenderer.Utils
{
    internal static class ColorUtils
    {
        internal static Color GetAdjustedColor(PathFillMode fillColorSource, Color fc)
        {
            switch (fillColorSource)
            {
                case PathFillMode.Darken:
                    fc = ApplyBlend(fc, Color.Black, 0.4);
                    break;
                case PathFillMode.DarkenLess:
                    fc = ApplyBlend(fc, Color.Black, 50D/255D);
                    break;
                case PathFillMode.LightenLess:
                    fc = ApplyBlend(fc, Color.White, 50D / 255D);
                    break;
                case PathFillMode.Lighten:
                    fc = ApplyBlend(fc, Color.White, 0.4);
                    break;
            }

            return fc;        
        }
        internal static Color ApplyBlend(Color color, Color blendColor, double percent)
        {
            var colorPercent = 1 - percent;
            var r = (int)Math.Min(255D, color.R * colorPercent + blendColor.R * percent);
            var g = (int)Math.Min(255D, color.G * colorPercent + blendColor.G * percent);
            var b = (int)Math.Min(255D, color.B * colorPercent + blendColor.B * percent);
            return Color.FromArgb(0xff, r, g, b);
        }
    }
}
