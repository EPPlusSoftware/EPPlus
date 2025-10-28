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
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusImageRenderer.Svg
{
    internal abstract class SvgChartObject : DrawingChart
    {
        internal SvgChartObject(ExcelChart chart) : base(chart)
        {
        }
        internal void SetMargins(ExcelTextBody tb)
        {
            tb.GetInsetsOrDefaults(out double l, out double r, out double t, out double b);
            LeftMargin = l.PointToPixel();
            RightMargin = r.PointToPixel();
            TopMargin = t.PointToPixel();
            BottomMargin = b.PointToPixel();
        }
        internal double LeftMargin { get; set; }
        internal double RightMargin { get; set; }
        internal double TopMargin { get; set; }
        internal double BottomMargin { get; set; }
        internal SvgRenderRectItem Rectangle { get; set; }
        public string Text { get; set; }
        protected static SvgRenderRectItem GetRectFromManualLayout(SvgChart sc, ExcelLayout layout)
        {
            var rect = new SvgRenderRectItem(sc.Chart);
            var ml = layout.ManualLayout;
            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.X = sc.Size.Width * (float)(layout.ManualLayout.Left ?? 0D) / 100;
                rect.Width = sc.Size.Width * (float)(layout.ManualLayout.Width ?? 0D) / 100;
            }
            else
            {
                //TODO:Add factor from default position
            }
            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Y = sc.Size.Height * (float)(layout.ManualLayout.Top ?? 0D) / 100;
                rect.Height = sc.Size.Height * (float)(layout.ManualLayout.Width ?? 0D) / 100;
            }
            else
            {
                //TODO:Add factor from default position
            }

            return rect;
        }
    }
}
