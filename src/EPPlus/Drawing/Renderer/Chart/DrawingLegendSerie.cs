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
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using System;

namespace EPPlusImageRenderer.Svg
{
    internal class DrawingLegendSerie : DrawingLegendSeriesIcon
    {
        internal DrawingTextbody Textbox { get; set; }

        internal void GetIconTopLeft(out double top, out double left)
        {
            if (SeriesIcon is LineRenderItem line)
            {
                top = line.Y1;
                left = line.X1;
            }
            else if (SeriesIcon is RectRenderItem rect)
            {
                top = rect.Top;
                left = rect.Left;
            }
            else
            {
                top = SeriesIcon.Bounds.Top;
                left = SeriesIcon.Bounds.Left;
            }
        }
    }
}