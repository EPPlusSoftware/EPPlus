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
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartPlotarea : SvgChartObject
    {
        public override RenderItemType Type => throw new System.NotImplementedException();

        public SvgChartPlotarea(SvgChart sc) : base(sc.Chart)
        {
            Rectangle = GetPlotAreaRectangle(sc);
        }
        internal SvgRenderRectItem GetPlotAreaRectangle(SvgChart sc)
        {
            var pa = sc.Chart.PlotArea;
            var rect = new SvgRenderRectItem(sc.Chart);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(sc, pa.Layout);
            }
            else
            {
                rect.Y = sc.Title.Rectangle?.Height ?? 0;
                rect.X = sc.VerticalAxis?.Rectangle.Width ?? 0;
                rect.Width = sc.Size.Width - rect.X;
                rect.Height = sc.Size.Height - rect.Y;

                switch(sc.Chart.Legend.Position)
                {
                    case eLegendPosition.Top:
                        break;
                }
            }

            return rect;
        }

        public override void Render(StringBuilder sb)
        {
            Rectangle.Render(sb);
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            throw new System.NotImplementedException();
        }
    }
}
