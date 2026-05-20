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
using EPPlus.Export.ImageRenderer.Svg.Chart;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartPlotareaRenderer : ChartDrawingObject
    {
        public ChartPlotareaRenderer(ChartRenderer sc) : base(sc)
        {
        }
        public List<ChartTypeDrawer> ChartTypeDrawers { get; set; } 
        internal void SetPlotAreaRectangle()
        {
            var pa = Chart.PlotArea;
            TopMargin = BottomMargin = LeftMargin = RightMargin = 10.5; //14px
            var rect = new RectRenderItem(ChartRenderer.Bounds);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(ChartRenderer, pa.Layout);
            }
            else
            {
                rect.Top = GetPlotAreaTop(ChartRenderer);
                rect.Left = GetPlotAreaLeft(ChartRenderer);
                rect.Width = GetPlotAreaWidth(ChartRenderer, rect);
                rect.Height = GetPlotAreaHeight(ChartRenderer, rect);
            }

            rect.SetDrawingPropertiesFill(pa.Fill, ChartRenderer.Chart.StyleManager.Style.PlotArea.FillReference.Color);
            rect.SetDrawingPropertiesBorder(pa.Border, ChartRenderer.Chart.StyleManager.Style.PlotArea.BorderReference.Color, pa.Border.Fill.Style != eFillStyle.NoFill, 0.75);
            Rectangle = rect;
        }

        private double GetPlotAreaHeight(SvgRenderRectItem rect)
        {
            var bottomAxis = GetAxisByPosition(eAxisPosition.Bottom);
            double vaHeight = 0;
            if (bottomAxis!=null)
            {
                 vaHeight = (bottomAxis.Rectangle?.Height ?? 0D) + (bottomAxis.Title?.TextBox?.GetActualHeight() ?? 0D);
            }
            if (sc.Chart.Legend?.Position == eLegendPosition.Bottom)
            {
                vaHeight += sc.Legend.Rectangle.Height + sc.Legend.TopMargin;
            }
            return sc.Bounds.Height - rect.GlobalTop - vaHeight - BottomMargin;
        }

        private double GetPlotAreaWidth(DrawingChart sc, SvgRenderRectItem rect)
        {
            var rightAxis = GetAxisByPosition(sc, eAxisPosition.Right);
            var lp = sc.Chart.Legend?.Position;
            var right = ((lp == eLegendPosition.Right || lp == eLegendPosition.TopRight) && sc.Legend != null ?
                        sc.Legend.Bounds.GlobalLeft - RightMargin :
                        sc.ChartArea.Rectangle.Width - RightMargin);


            double rightAxisWidth;
            if (rightAxis == null)
            {
                rightAxisWidth =  0;
            }
            else
            {
                rightAxisWidth = (rightAxis.Title?.TextBox.GetActualWidth() ?? 0D) + (rightAxis.Rectangle?.Width ?? 0D);
            }

            var width = right - rightAxisWidth - rect.GlobalLeft;
            //Reserve space for the last label that will be on the tick label instead of Middle of the category.
            if (sc.HorizontalAxis != null && sc.VerticalAxis.Axis.CrossBetween == eCrossBetween.MidCat)
            {
                var minusPA = width / sc.HorizontalAxis.AxisValues.Count / 2;
                if (minusPA > rightAxisWidth)
                {
                    rightAxisWidth = minusPA;
                }
            }
            if (sc.SecondHorizontalAxis != null && sc.SecondVerticalAxis.Axis.CrossBetween == eCrossBetween.MidCat)
            {
                var minusSA = width / sc.SecondHorizontalAxis.AxisValues.Count / 2;
                if (minusSA > rightAxisWidth)
                {
                    rightAxisWidth = minusSA;
                }
            }

            return right - rightAxisWidth-rect.GlobalLeft;
        }
        private double GetPlotAreaLeft(DrawingChart sc)
        {
            var left = LeftMargin;
            if(sc.Chart.Legend?.Position == eLegendPosition.Left)
            {
                left += sc.Legend.Bounds.Width + sc.Legend.RightMargin;
            }

            var leftAxis = GetAxisByPosition(sc, eAxisPosition.Left);
            if (leftAxis != null)
            {
                if(leftAxis.Title!=null)
                {
                    left += leftAxis.Title.TextBox.GetActualWidth();
                }
                if (leftAxis.Rectangle != null)
                {
                    left += leftAxis.Rectangle.Width + 1.5;
                }
            }
            return left;
        }
        private double GetPlotAreaTop()
        {
            double haHeight = 0;
            var topAxis = GetAxisByPosition(eAxisPosition.Top);
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

            return (Chart.Legend?.Position == eLegendPosition.Top ? ChartRenderer.Legend.Bounds.Bottom : ChartRenderer.Title?.Rectangle?.GlobalBottom ?? 0d) + haHeight + TopMargin;
        }

        private SvgChartAxis GetAxisByPosition(eAxisPosition pos)
        {
            if (ChartRenderer.HorizontalAxis != null && ChartRenderer.HorizontalAxis.Axis.AxisPosition == pos)
            {
                return ChartRenderer.HorizontalAxis;
            }
            else if (ChartRenderer.VerticalAxis != null && ChartRenderer.VerticalAxis.Axis.AxisPosition == pos)
            {
                return ChartRenderer.VerticalAxis;
            }
            else if (ChartRenderer.SecondHorizontalAxis != null && ChartRenderer.SecondHorizontalAxis.Axis.AxisPosition == pos)
            {
                return ChartRenderer.SecondHorizontalAxis;
            }
            else if (ChartRenderer.SecondVerticalAxis != null && ChartRenderer.SecondVerticalAxis.Axis.AxisPosition == pos)
            {
                return ChartRenderer.SecondVerticalAxis;
            }
            return null;
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Rectangle);
        }

        internal void DrawSeries()
        {
            foreach (var drawer in ChartTypeDrawers)
            {
                drawer.DrawSeries();
            }
        }
    }
}
