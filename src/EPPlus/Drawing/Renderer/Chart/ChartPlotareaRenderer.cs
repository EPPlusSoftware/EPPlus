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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlusImageRenderer.Svg
{
    internal class ChartPlotareaRenderer : ChartDrawingObject
    {
        public ChartPlotareaRenderer(ChartRenderer sc) : base(sc)
        {
             
        }
        public List<ChartTypeDrawer> ChartTypeDrawers { get; set; } 
        public GroupRenderItem Group { get; private set; }
        internal void SetPlotAreaRectangle()
        {
            var pa = Chart.PlotArea;
            TopMargin = BottomMargin = LeftMargin = RightMargin = 10.5; //14px
            var rect = new RectRenderItem(CRenderer.Bounds);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(CRenderer, pa.Layout);
            }
            else
            {
                rect.Top = GetPlotAreaTop();
                rect.Left = GetPlotAreaLeft();
                rect.Width = GetPlotAreaWidth(rect);
                rect.Height = GetPlotAreaHeight(rect);
            }

            Group = new GroupRenderItem(CRenderer.Bounds);
            Group.Bounds.Top = rect.Top;
            Group.Bounds.Left = rect.Left;           
            rect.Top = rect.Left = 0;
            Group.RenderItems.Add(rect);

            if(CRenderer.Legend!=null && Chart.Legend.Position == eLegendPosition.Right ||
               Chart.Legend.Position == eLegendPosition.Left)
            {
                CRenderer.Legend.Rectangle.Top = Group.Top + rect.Height / 2 - CRenderer.Legend.Rectangle.Height / 2;
            }

            rect.SetDrawingPropertiesFill(CRenderer.Theme, pa.Fill, CRenderer.Chart.StyleManager.Style?.PlotArea.FillReference.Color, false, DefaultFillColor);
            rect.SetDrawingPropertiesBorder(CRenderer.Theme, pa.Border, CRenderer.Chart.StyleManager.Style?.PlotArea.BorderReference.Color, pa.Border.Fill.Style != eFillStyle.NoFill, DefaultBorderColor, 0.75);
            Rectangle = rect;
        }

        private double GetPlotAreaHeight(RectRenderItem rect)
        {
            var bottomAxis = GetAxisActualByPosition(eActualAxisPosition.Bottom);
            double vaHeight = 0;
            if (bottomAxis!=null)
            {
                var bottomSecondAxis = GetAxisActualByPosition(eActualAxisPosition.BottomSecond);
                vaHeight = (bottomAxis.Rectangle?.Height ?? 0D) + (bottomAxis.Title?.TextBox?.GetActualHeight() ?? 0D) + (bottomSecondAxis?.Rectangle?.Height ?? 0D);
            }
            if (Chart.Legend?.Position == eLegendPosition.Bottom)
            {
                vaHeight += CRenderer.Legend.Rectangle.Height + CRenderer.Legend.TopMargin;
            }
            return CRenderer.Bounds.Height - rect.GlobalTop - vaHeight - BottomMargin;
        }

        private double GetPlotAreaWidth(RectRenderItem rect)
        {
            var rightAxis = GetAxisActualByPosition(eActualAxisPosition.Right);
            var rightSecondAxis = GetAxisActualByPosition(eActualAxisPosition.RightSecond);
            var lp = CRenderer.Chart.Legend?.Position;
            var right = ((lp == eLegendPosition.Right || lp == eLegendPosition.TopRight) && CRenderer.Legend != null ?
                        CRenderer.Legend.Rectangle.Bounds.GlobalLeft - RightMargin :
                        CRenderer.ChartArea.Rectangle.Width - RightMargin);


            double rightAxisWidth;
            if (rightAxis == null)
            {
                rightAxisWidth =  0;
            }
            else
            {
                rightAxisWidth = (rightAxis.Title?.TextBox.GetActualWidth() ?? 0D) + (rightAxis.Rectangle?.Width ?? 0D) + (rightSecondAxis?.Rectangle?.Width ?? 0D);
            }

            var width = right - rightAxisWidth - rect.GlobalLeft;
            //Reserve space for the last label that will be on the tick label instead of Middle of the category.
            if (CRenderer.HorizontalAxis != null && CRenderer.VerticalAxis.Axis.CrossBetween == eCrossBetween.MidCat)
            {
                var minusPA = width / CRenderer.HorizontalAxis.AxisValues.Count / 2;
                if (minusPA > rightAxisWidth)
                {
                    rightAxisWidth = minusPA;
                }
            }
            if (CRenderer.SecondHorizontalAxis != null && CRenderer.SecondVerticalAxis.Axis.CrossBetween == eCrossBetween.MidCat)
            {
                var minusSA = width / CRenderer.SecondHorizontalAxis.AxisValues.Count / 2;
                if (minusSA > rightAxisWidth)
                {
                    rightAxisWidth = minusSA;
                }
            }

            return right - rightAxisWidth-rect.GlobalLeft;
        }
        private double GetPlotAreaLeft()
        {
            var left = LeftMargin;
            if(CRenderer.Chart.Legend?.Position == eLegendPosition.Left)
            {
                left += CRenderer.Legend.Rectangle.Bounds.Width + CRenderer.Legend.RightMargin;
            }

            var leftAxis = GetAxisActualByPosition(eActualAxisPosition.Left);
            var leftSecondAxis = GetAxisActualByPosition(eActualAxisPosition.LeftSecond);
            if (leftAxis == null)
            {
                leftAxis = GetAxisByPosition(eAxisPosition.Left);
                left += leftAxis?.Title?.TextBox.Width ?? 0D;
            }
            else
            {
                if (leftAxis.Title!=null)
                {
                    left += leftAxis.Title.TextBox.GetActualWidth();
                }
                if (leftAxis.Rectangle != null)
                {
                    left += leftAxis.Rectangle.Width + 1.5;
                }
                if(leftSecondAxis!=null)
                {
                    left += leftSecondAxis.Rectangle.Width;
                }
            }
            return left;
        }
        private double GetPlotAreaTop()
        {
            double haHeight = 0;
            var topAxis = GetAxisActualByPosition(eActualAxisPosition.Top);
            var topSecondAxis = GetAxisActualByPosition(eActualAxisPosition.TopSecond);
            if (topAxis == null)
            {
                //If the axis is not on the top, we should check if there is an axis that has the position on the top. If there is, we should reserve space for the title of the axis. This can happen when LabelPosition is set to Low and the axis is on the bottom, but the position of the axis is set to top.
                topAxis = GetAxisByPosition(eAxisPosition.Top);
                haHeight = topAxis?.Title?.Rectangle.Height ?? 0D;
            }
            else
            {
                haHeight = (topAxis.Rectangle?.Height ?? 0D) + (topSecondAxis?.Rectangle?.Height ?? 0D) + (topAxis.Title?.TextBox?.GetActualHeight() ?? 0D);
            }

            return (Chart.Legend?.Position == eLegendPosition.Top ? CRenderer.Legend.Rectangle.Bounds.Bottom : CRenderer.Title?.Rectangle?.GlobalBottom ?? 0d) + haHeight + TopMargin;
        }

        private ChartAxisRenderer GetAxisActualByPosition(eActualAxisPosition pos)
        {
            if (CRenderer.HorizontalAxis != null && CRenderer.HorizontalAxis.Axis.ActualAxisPosition == pos)
            {
                return CRenderer.HorizontalAxis;
            }
            else if (CRenderer.VerticalAxis != null && CRenderer.VerticalAxis.Axis.ActualAxisPosition == pos)
            {
                return CRenderer.VerticalAxis;
            }
            else if (CRenderer.SecondHorizontalAxis != null && CRenderer.SecondHorizontalAxis.Axis.ActualAxisPosition == pos)
            {
                return CRenderer.SecondHorizontalAxis;
            }
            else if (CRenderer.SecondVerticalAxis != null && CRenderer.SecondVerticalAxis.Axis.ActualAxisPosition == pos)
            {
                return CRenderer.SecondVerticalAxis;
            }
            return null;
        }
        private ChartAxisRenderer GetAxisByPosition(eAxisPosition pos)
        {
            if (CRenderer.HorizontalAxis != null && CRenderer.HorizontalAxis.Axis.AxisPosition == pos)
            {
                return CRenderer.HorizontalAxis;
            }
            else if (CRenderer.VerticalAxis != null && CRenderer.VerticalAxis.Axis.AxisPosition == pos)
            {
                return CRenderer.VerticalAxis;
            }
            else if (CRenderer.SecondHorizontalAxis != null && CRenderer.SecondHorizontalAxis.Axis.AxisPosition == pos)
            {
                return CRenderer.SecondHorizontalAxis;
            }
            else if (CRenderer.SecondVerticalAxis != null && CRenderer.SecondVerticalAxis.Axis.AxisPosition == pos)
            {
                return CRenderer.SecondVerticalAxis;
            }
            return null;
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.Add(Group);
        }

        internal void DrawSeries()
        {
            foreach (var drawer in ChartTypeDrawers)
            {
                drawer.DrawSeries();
            }
        }
        internal override Color? DefaultFillColor { get => null;  }
    }
}
