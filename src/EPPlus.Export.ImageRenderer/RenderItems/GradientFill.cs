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
using EPPlus.Export.ImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using System.Collections.Generic;
using System.Drawing;
using EPPlusColorConverter = OfficeOpenXml.Utils.TypeConversion.ColorConverter;
using System;

namespace EPPlusImageRenderer.RenderItems
{
    internal class DrawGradientFill
    {
        public DrawGradientFill(ExcelTheme theme, ExcelDrawingGradientFill gradientFill)
        {
            this.Settings = gradientFill;
            for (int i = 0; i < gradientFill.Colors.Count; i++)
            {
                var c = new GradientFillColor(gradientFill.Colors[i].Position, EPPlusColorConverter.GetThemeColor(theme, gradientFill.Colors[i].Color));
                Colors.Add(c);
            }

        }

        public DrawGradientFill(List<Color> colors, List<double> stops)
        {
            for (int i = 0; i < stops.Count; i++)
            {
                var c = new GradientFillColor(stops[i], colors[i]);
                Colors.Add(c);
            }
        }

        public ExcelDrawingGradientFill Settings { get; set; }
        public List<GradientFillColor> Colors { get; set; } = new List<GradientFillColor>();
    }
}