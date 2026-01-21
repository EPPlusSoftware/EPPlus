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
using System.Collections.Generic;

namespace EPPlusImageRenderer.Svg
{
    internal abstract class SvgChartObject 
    {
        internal ExcelChart Chart { get; }
        internal SvgChartObject(ExcelChart chart) 
        {
            Chart= chart; 
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
        internal SvgRenderLineItem Line { get; set; }
        public string Text { get; set; }
        protected static SvgRenderRectItem GetRectFromManualLayout(SvgChart sc, ExcelLayout layout)
        {
            var rect = new SvgRenderRectItem(sc.Chart);
            var ml = layout.ManualLayout;
            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Left = sc.Bounds.Width * (float)(layout.ManualLayout.Left ?? 0D) / 100;
            }
            else
            {
                //TODO:Add factor from default position
            }
            //Width is always factor.
            rect.Width = sc.Bounds.Width * ml.GetWidth() / 100;

            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Top = sc.Bounds.Height * (float)(layout.ManualLayout.Top ?? 0D) / 100;
            }
            else
            {
                //TODO:Add factor from default position
            }
            //Height is always factor.
            rect.Height = sc.Bounds.Height * ml.GetHeight() / 100;
            return rect;
        }
        internal abstract void AppendRenderItems(List<RenderItemBase> renderItems);

    }
}
