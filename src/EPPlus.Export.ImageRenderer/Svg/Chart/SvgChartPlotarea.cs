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
            TopMargin = BottomMargin = LeftMargin = RightMargin = 10.5; //14px
            var rect = new SvgRenderRectItem(sc, sc.Bounds);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(sc, pa.Layout);
            }
            else
            {
                rect.Top = GetPlotAreaTop(sc);
                rect.Left = GetPlotAreaLeft(sc);
                rect.Width = GetPlotAreaWidth(sc, rect);
                rect.Height = GetPlotAreaHeight(sc, rect);
            }

            rect.SetDrawingPropertiesFill(pa.Fill, sc.Chart.StyleManager.Style.PlotArea.FillReference.Color);
            rect.SetDrawingPropertiesBorder(pa.Border, sc.Chart.StyleManager.Style.PlotArea.BorderReference.Color, pa.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            return rect;
        }

        private double GetPlotAreaHeight(SvgChart sc, SvgRenderRectItem rect)
        {
            var bottomAxis = GetAxisByPosition(sc, eAxisPosition.Bottom);
            double vaHeight = 0, vaTitleHeight = 0;
            if (bottomAxis!=null)
            {
                 vaHeight = (bottomAxis.Rectangle?.Height ?? 0D) + (bottomAxis.Title?.TextBox?.GetActualHeight() ?? 0D);
            }
            if (sc.Chart.Legend?.Position == eLegendPosition.Bottom)
            {
                vaHeight += sc.Legend.Rectangle.Height;
            }
            return sc.Bounds.Height - rect.GlobalTop - vaHeight - vaTitleHeight - BottomMargin;
        }

        private double GetPlotAreaWidth(SvgChart sc, SvgRenderRectItem rect)
        {
            var rightAxis = GetAxisByPosition(sc, eAxisPosition.Right);
            var lp = sc.Chart.Legend?.Position;
            var left = ((lp == eLegendPosition.Right || lp == eLegendPosition.TopRight) && sc.Legend != null ?
                        sc.Legend.Bounds.GlobalLeft - RightMargin :
                        sc.ChartArea.Rectangle.Width - RightMargin);
            if (rightAxis == null)
            {
                return left - rect.GlobalLeft;
            }
            else
            {
                var rightWidth = rightAxis.Title?.TextBox.GetActualWidth() ?? 0D + rightAxis.Rectangle?.Width ?? 0D;
                return left - rightWidth - rect.Left;
            }
        }
        private double GetPlotAreaLeft(SvgChart sc)
        {
            var leftAxis = GetAxisByPosition(sc, eAxisPosition.Left);
            if (leftAxis == null || (leftAxis.Rectangle==null && leftAxis.Title?.Rectangle==null))
            {
                return sc.Chart.Legend?.Position == eLegendPosition.Left ? sc.Legend.Bounds.Right + LeftMargin : LeftMargin;
            }
            else
            {
                return leftAxis.Rectangle == null ? leftAxis.Title.TextBox.GetActualWidth(): leftAxis.Rectangle.GlobalRight;
            }
        }

        private double GetPlotAreaTop(SvgChart sc)
        {
            double haHeight = 0;
            var topAxis = GetAxisByPosition(sc, eAxisPosition.Top);
            if (topAxis == null)
            {
                //var bottomAxis = GetAxisByPosition(sc, eAxisPosition.Bottom);
                //if (bottomAxis != null && sc.Chart.XAxis.LabelPosition == eTickLabelPosition.High)
                //{
                //    top += bottomAxis.Rectangle.Height;
                //}
                //return top;
            }
            else
            {
                haHeight = (topAxis.Rectangle?.Height ?? 0D) + (topAxis.Title?.TextBox?.GetActualHeight() ?? 0D);
            }

            return (sc.Chart.Legend?.Position == eLegendPosition.Top ? sc.Legend.Bounds.Bottom : sc.Title?.Rectangle?.GlobalBottom ?? 0d) + haHeight + TopMargin;
        }

        private SvgChartAxis GetAxisByPosition(SvgChart sc, eAxisPosition pos)
        {
            if (sc.HorizontalAxis != null && sc.HorizontalAxis.Axis.AxisPosition == pos)
            {
                return sc.HorizontalAxis;
            }
            else if (sc.VerticalAxis != null && sc.VerticalAxis.Axis.AxisPosition == pos)
            {
                return sc.VerticalAxis;
            }
            else if (sc.SecondHorizontalAxis != null && sc.SecondHorizontalAxis.Axis.AxisPosition == pos)
            {
                return sc.SecondHorizontalAxis;
            }
            else if (sc.SecondVerticalAxis != null && sc.SecondVerticalAxis.Axis.AxisPosition == pos)
            {
                return sc.SecondVerticalAxis;
            }
            return null;
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }
    }
}
