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
using OfficeOpenXml.Drawing;
using System.Drawing;
using TC = OfficeOpenXml.Utils.TypeConversion;

namespace EPPlusImageRenderer.Utils
{
    internal static class ColorUtils
    {
        internal static Color GetAdjustedColor(PathFillMode fillColorSource, Color fc)
        {
            switch (fillColorSource)
            {
                case PathFillMode.Darken:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.Black, 0.4);
                    break;
                case PathFillMode.DarkenLess:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.Black, 50D/255D);
                    break;
                case PathFillMode.LightenLess:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.White, 50D / 255D);
                    break;
                case PathFillMode.Lighten:
                    fc = TC.ColorConverter.ApplyBlend(fc, Color.White, 0.4);
                    break;
            }

            return fc;
        }
    }
}
