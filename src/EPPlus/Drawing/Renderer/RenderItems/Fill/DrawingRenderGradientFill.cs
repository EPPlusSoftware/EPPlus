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
using EPPlus.Export.Utils;

namespace EPPlus.DrawingRenderer.RenderItems
{
    internal class DrawingRenderGradientFill : RenderGradientFill
    {
        public DrawingRenderGradientFill(ExcelTheme theme, ExcelDrawingGradientFill gradientFill) : base()
        {
            //this.Settings = gradientFill;
            for (int i = 0; i < gradientFill.Colors.Count; i++)
            {
                var opacity = EPPlusColorConverter.GetOpacity(gradientFill.Colors[i].Color);
                var c = new GradientFillColor(gradientFill.Colors[i].Position, EPPlusColorConverter.GetThemeColor(theme, gradientFill.Colors[i].Color));
                c.Opacity = opacity;
                Colors.Add(c);
            }
            FocusPoint = gradientFill.FocusPoint.AsOffsetRectangle();
            TileRectangle = gradientFill.TileRectangle.AsOffsetRectangle();
            ShadePath = (ShadePath)gradientFill.ShadePath;
            LinearSettings.Angle = gradientFill.LinearSettings.Angle;
            LinearSettings.Scaled = gradientFill.LinearSettings.Scaled;
        }


        //public DrawingRenderGradientFill(List<Color> colors, List<double> stops) : base(colors, stops)
        //{
        //}
    }
}