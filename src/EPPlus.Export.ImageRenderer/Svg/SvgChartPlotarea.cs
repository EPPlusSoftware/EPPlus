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
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartPlotarea : SvgChartObject
    {
        public SvgChartPlotarea(SvgChart sc) : base(sc.Chart)
        {
            SvgChart = sc;
            Rectangle = GetPlotAreaRectangle(sc);
        }
        public SvgChart SvgChart { get; set; }
        internal SvgRenderRectItem GetPlotAreaRectangle(SvgChart sc)
        {
            var pa = sc.Chart.PlotArea;
            TopMargin = BottomMargin = LeftMargin = RightMargin = 14;
            var rect = new SvgRenderRectItem(sc.Chart);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(sc, pa.Layout);
            }
            else
            {
                var lp = sc.Chart.Legend.Position;
                rect.Top = (sc.Title?.Rectangle?.Bottom ?? 0d) + TopMargin;
                rect.Left = lp == eLegendPosition.Left ? sc.Legend.Rectangle.Right + LeftMargin : LeftMargin;
                rect.Width = lp == eLegendPosition.Right || lp == eLegendPosition.TopRight ? sc.Legend.Rectangle.Left - RightMargin : RightMargin;
                rect.Height = sc.Size.Height - rect.Top;
            }

            return rect;
        }

        public override void Render(StringBuilder sb)
        {
            Rectangle.Render(sb);
        }
    }
}
