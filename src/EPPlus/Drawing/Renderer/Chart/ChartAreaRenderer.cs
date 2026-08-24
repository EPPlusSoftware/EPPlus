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
using EPPlus.DrawingRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Renderer.Chart.ChartElementStyleTables;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAreaRenderer : ChartDrawingObjectWithDefaults
    {
        public ChartAreaRenderer(ChartRenderer sc, SvgRenderOptions options) : base(sc)
        {
            if(options.Size.Width.HasValue)
            {
                sc.Bounds.Width = options.Size.WidthPixels;
            }
            if (options.Size.Height.HasValue)
            {
                sc.Bounds.Height = options.Size.HeightPixels;
            }

            Rectangle = new RectRenderItem(sc.Bounds);
        }

        internal override Color? DefaultFillColor { get => ChartRenderer.Theme.ColorScheme.Light1.GetColor(); }
        internal override Color? DefaultBorderColor
        {
            get
            {
                return Color.FromArgb(0x89, 0x89, 0x89);
            }
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }

        internal override Color? GetDefaultBorderColor()
        {
            //We only get here if the node is null or empty
            var themedLine = GetThemedLine(ChartElement.ChartArea, (int)Chart.Style, Chart.Border.Fill != null && Chart.Border.Fill.IsEmpty, out Color? lineColor);
            ////Kept here in case needed in future for effect etc.
            //var themedLine = GetThemedLine(ChartElement.ChartArea, (int)Chart.Style, out Color? lineCol);
            return lineColor;
        }

        internal override Color? GetDefaultFillColor()
        {
            return GetDefaultFillColorForElement(ChartElement.ChartArea, (int)Chart.Style);
        }
    }
}
