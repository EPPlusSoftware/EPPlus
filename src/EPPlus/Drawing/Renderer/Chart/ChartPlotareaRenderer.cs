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
            var rect = new RectRenderItem(ChartRenderer.Bounds);
            if (pa.Layout.HasLayout)
            {
                rect = GetRectFromManualLayout(ChartRenderer, pa.Layout);
            }
            else
            {
                rect.Top = GetPlotAreaTop();
                rect.Left = GetPlotAreaLeft();
                rect.Width = GetPlotAreaWidth(rect);
                rect.Height = GetPlotAreaHeight(rect);
            }

            Group = new GroupRenderItem(ChartRenderer.Bounds);
            Group.Bounds.Top = rect.Top;
            Group.Bounds.Left = rect.Left;           
            rect.Top = rect.Left = 0;
            Group.RenderItems.Add(rect);

            if(ChartRenderer.Legend!=null && Chart.Legend.Position == eLegendPosition.Right ||
               Chart.Legend.Position == eLegendPosition.Left)
            {
                ChartRenderer.Legend.Rectangle.Top = Group.Top + rect.Height / 2 - ChartRenderer.Legend.Rectangle.Height / 2;
            }

            rect.SetDrawingPropertiesFill(ChartRenderer.Theme, pa.Fill, ChartRenderer.Chart.StyleManager.Style?.PlotArea.FillReference.Color, UserSpaceSettings.ObjectBoundingBox, DefaultFillColor);
            rect.SetDrawingPropertiesBorder(ChartRenderer.Theme, pa.Border, ChartRenderer.Chart.StyleManager.Style?.PlotArea.BorderReference.Color, pa.Border.Fill.Style != eFillStyle.NoFill, DefaultBorderColor, 0.75);
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
            else
            {
                var bottomAx = GetAxisByPosition(eAxisPosition.Bottom);
                if(bottomAx!=null) //Title is always placed on bottom.
                {
                    vaHeight = bottomAx.Title?.TextBox?.GetActualHeight() ?? 0D;
                }
            }
            if (Chart.Legend?.Position == eLegendPosition.Bottom)
            {
                vaHeight += ChartRenderer.Legend.Rectangle.Height + ChartRenderer.Legend.TopMargin;
            }
            return ChartRenderer.Bounds.Height - rect.GlobalTop - vaHeight - BottomMargin;
        }

        private double GetPlotAreaWidth(RectRenderItem rect)
        {
            var rightActualAxis = GetAxisActualByPosition(eActualAxisPosition.Right);
            var rightSecondAxis = GetAxisActualByPosition(eActualAxisPosition.RightSecond);
            var lp = ChartRenderer.Chart.Legend?.Position;
            var right = ((lp == eLegendPosition.Right || lp == eLegendPosition.TopRight) && ChartRenderer.Legend != null ?
                        ChartRenderer.Legend.Rectangle.Bounds.GlobalLeft - RightMargin :
                        ChartRenderer.ChartArea.Rectangle.Width - RightMargin);


            double rightAxisWidth;
            if (rightActualAxis == null)
            {
                var rightAxis = GetAxisByPosition(eAxisPosition.Right);
                if (rightAxis == null)
                {
                    rightAxisWidth = 0;
                }
                else
                {
                    rightAxisWidth = rightAxis.Title?.TextBox.GetActualWidth() ?? 0D;
                }
            }
            else
            {
                rightAxisWidth = (rightActualAxis.Title?.TextBox.GetActualWidth() ?? 0D) + (rightActualAxis.Rectangle?.Width ?? 0D) + (rightSecondAxis?.Rectangle?.Width ?? 0D);
            }

            var width = right - rightAxisWidth - rect.GlobalLeft;
            //Reserve space for the last label that will be on the tick label instead of Middle of the category.
            if (ChartRenderer.HorizontalAxis != null && ChartRenderer.VerticalAxis.Axis.CrossBetween == eCrossBetween.MidCat)
            {
                var minusPA = width / ChartRenderer.HorizontalAxis.AxisValues.Count / 2;
                if (minusPA > rightAxisWidth)
                {
                    rightAxisWidth = minusPA;
                }
            }
            if (ChartRenderer.SecondHorizontalAxis != null && ChartRenderer.SecondVerticalAxis.Axis.CrossBetween == eCrossBetween.MidCat)
            {
                var minusSA = width / ChartRenderer.SecondHorizontalAxis.AxisValues.Count / 2;
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
            if(ChartRenderer.Chart.Legend?.Position == eLegendPosition.Left)
            {
                left += ChartRenderer.Legend.Rectangle.Bounds.Width + ChartRenderer.Legend.RightMargin;
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

            return (Chart.Legend?.Position == eLegendPosition.Top ? ChartRenderer.Legend.Rectangle.Bounds.Bottom : ChartRenderer.Title?.Rectangle?.GlobalBottom ?? 0d) + haHeight;
        }

        private ChartAxisRenderer GetAxisActualByPosition(eActualAxisPosition pos)
        {
            if (ChartRenderer.HorizontalAxis != null && ChartRenderer.HorizontalAxis.Axis.ActualAxisPosition == pos)
            {
                return ChartRenderer.HorizontalAxis;
            }
            else if (ChartRenderer.VerticalAxis != null && ChartRenderer.VerticalAxis.Axis.ActualAxisPosition == pos)
            {
                return ChartRenderer.VerticalAxis;
            }
            else if (ChartRenderer.SecondHorizontalAxis != null && ChartRenderer.SecondHorizontalAxis.Axis.ActualAxisPosition == pos)
            {
                return ChartRenderer.SecondHorizontalAxis;
            }
            else if (ChartRenderer.SecondVerticalAxis != null && ChartRenderer.SecondVerticalAxis.Axis.ActualAxisPosition == pos)
            {
                return ChartRenderer.SecondVerticalAxis;
            }
            return null;
        }
        private ChartAxisRenderer GetAxisByPosition(eAxisPosition pos)
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
