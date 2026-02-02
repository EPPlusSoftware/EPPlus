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
using EPPlus.Export.ImageRenderer.Svg.Chart;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartPlotarea : SvgChartObject
    {
        public SvgChartPlotarea(SvgChart sc) : base(sc)
        {
            SvgChart = sc;
            Rectangle = GetPlotAreaRectangle(sc);
        }
        public SvgChart SvgChart { get; set; }
        public List<ChartTypeDrawer> ChartTypeDrawers { get; set; } 
        internal SvgRenderRectItem GetPlotAreaRectangle(SvgChart sc)
        {
            var pa = sc.Chart.PlotArea;
            TopMargin = BottomMargin = LeftMargin = RightMargin = 14;
            var rect = new SvgRenderRectItem(sc, sc.Bounds);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(sc, pa.Layout);
            }
            else
            {
                var lp = sc.Chart.Legend?.Position;
                rect.Top = (lp==eLegendPosition.Top ? sc.Legend.Rectangle.GlobalBottom : sc.Title?.Rectangle?.GlobalBottom ?? 0d) + TopMargin;
                if(sc.HorizontalAxis!=null && sc.Chart.XAxis.LabelPosition==eTickLabelPosition.High)
                {
                    rect.Top += sc.HorizontalAxis.Rectangle.Height;
                }
                rect.Left = lp == eLegendPosition.Left ? sc.Legend.Rectangle.GlobalRight + LeftMargin : LeftMargin;                
                if(sc.VerticalAxis!=null)
                {
                    rect.Left = sc.VerticalAxis.Rectangle?.GlobalRight ?? sc.VerticalAxis.Title.Rectangle.GlobalRight;
                }

                rect.Width = (lp == eLegendPosition.Right || lp == eLegendPosition.TopRight ? 
                        sc.Legend.Bounds.GlobalLeft - RightMargin : 
                        sc.ChartArea.Width - RightMargin) 
                  - rect.GlobalLeft;

                double vaHeight=0, vaTitleHeight=0;
                if(sc.HorizontalAxis != null)
                {
                    vaHeight = (sc.HorizontalAxis.Rectangle?.Height ?? 0D) + (sc.HorizontalAxis.Title?.Rectangle?.Height ?? 0D);
                }
                if(lp==eLegendPosition.Bottom)
                {
                    vaHeight += sc.Legend.Rectangle.Height;
                }
                rect.Height = sc.Bounds.Height - rect.GlobalTop - vaHeight - vaTitleHeight - BottomMargin;                
            }

            rect.SetDrawingPropertiesFill(pa.Fill, sc.Chart.StyleManager.Style.PlotArea.FillReference.Color);
            rect.SetDrawingPropertiesBorder(pa.Border, sc.Chart.StyleManager.Style.PlotArea.BorderReference.Color, pa.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            return rect;
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }

    }
}
