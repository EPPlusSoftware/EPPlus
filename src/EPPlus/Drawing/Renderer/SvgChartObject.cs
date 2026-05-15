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
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Collections.Generic;

namespace EPPlusImageRenderer.Svg
{
    public abstract class DrawingObject
    {
        internal virtual BoundingBox Bounds { get; set; }

        protected DrawingObject(BoundingBox parent)
        {
            DrawingRenderer = renderer;
            Bounds = new BoundingBox() { Parent=parent};
        }
        internal abstract void AppendRenderItems(List<RenderItem> renderItems);
    }
    internal abstract class DrawingObjectNoBounds
    {
        internal protected DrawingBase DrawingRenderer { get; }

        protected DrawingObjectNoBounds(DrawingBase renderer)
        {
            DrawingRenderer = renderer;
        }
        internal abstract void AppendRenderItems(List<RenderItem> renderItems);
    }
    internal class SvgChartArea : SvgChartObject
    {
        public SvgChartArea(SvgChart sc) : base(sc)
        {
            Rectangle = new SvgRenderRectItem(sc, sc.Bounds);
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }
    }
    internal abstract class SvgChartObject : DrawingObject
    {
        internal DrawingChart ChartRenderer;
        internal ExcelChart Chart => (ExcelChart)ChartRenderer.Drawing;
        internal SvgChartObject(DrawingChart chart) : base(chart, chart.Bounds)
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
        internal SvgRenderRectItem Rectangle { get; set; }
        protected static SvgRenderRectItem GetRectFromManualLayout(SvgChart sc, ExcelLayout layout, BoundingBox parent=null)
        {
            var bounds = parent ?? sc.Bounds;
            var rect = new SvgRenderRectItem(sc, sc.ChartArea.Rectangle.Bounds);
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
