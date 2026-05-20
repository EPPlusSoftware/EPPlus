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
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Collections.Generic;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartAreaRenderer : ChartDrawingObject
    {
        public ChartAreaRenderer(ChartRenderer sc) : base(sc)
        {
            Rectangle = new RectRenderItem(sc.Bounds);
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }
    }
    internal abstract class ChartDrawingObject : DrawingObject
    {
        internal ChartRenderer ChartRenderer;
        internal ExcelChart Chart => (ExcelChart)ChartRenderer.Drawing;
        internal ChartDrawingObject(ChartRenderer chart)
        {
            ChartRenderer = chart;
        }
        internal void SetMargins(ExcelTextBody tb)
        {
            tb.GetInsetsOrDefaults(out double l, out double r, out double t, out double b);
            LeftMargin = l;
            RightMargin = r;
            TopMargin = t;
            BottomMargin = b;
        }
        internal double LeftMargin { get; set; }
        internal double RightMargin { get; set; }
        internal double TopMargin { get; set; }
        internal double BottomMargin { get; set; }
        internal RectRenderItem Rectangle { get; set; }
        protected static RectRenderItem GetRectFromManualLayout(ChartRenderer sc, ExcelLayout layout, BoundingBox parent=null)
        {
            var bounds = parent ?? sc.Bounds;
            var rect = new RectRenderItem(sc.ChartArea.Rectangle.Bounds);
            var ml = layout.ManualLayout;
            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Left = bounds.Width * (float)(layout.ManualLayout.Left ?? 0D) / 100;
            }
            else
            {
                rect.Left = bounds.Width * (float)(ml.Left ?? 0D) / 100;
                //TODO:Add factor from default position
            }

            //Width is always factor.
            rect.Width = bounds.Width * ml.GetWidth() / 100;

            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Top = bounds.Height * (float)(layout.ManualLayout.Top ?? 0D) / 100;
            }
            else
            {
                rect.Top = bounds.Height * (float)(ml.Top ?? 0D) / 100;
                //TODO:Add factor from default position
            }
            //Height is always factor.
            rect.Height = bounds.Height * ml.GetHeight() / 100;
            return rect;
        }
    }
}
