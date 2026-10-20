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
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.ImageRenderer.Svg.Chart;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart.Defaults;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Drawing;
using d = OfficeOpenXml.Drawing.Renderer;
namespace EPPlusImageRenderer
{
    internal class ChartRenderer : d.DrawingRenderer
    {
        internal Color? GetDefaultFontColor()
        {
            var chartId = (int)Chart.Style;

            if(chartId == 209)
            {
                chartId = 102;
            }

            //In modern files the styles exist at ids 101->148 
            if(chartId > 100 && chartId < 148)
            {
                chartId = chartId - 100;
            }

            //Support for old chart styles 1-48
            //Chart styles  40 -> 48 are black background with white text
            if (chartId > 0 && chartId < 41)
            {
                return Color.Black;
            }
            else if(chartId> 41 && chartId < 49)
            {
                return Color.White;
            }

            //We only get here if the node is null or empty
            var themedLine = ChartDrawingObject.GetThemedLine(Theme, ChartElement.MajorGridLines, (int)Chart.Style, false, out Color? lineColor);
            ////Kept here in case needed in future for effect etc.
            //var themedLine = GetThemedLine(ChartElement.ChartArea, (int)Chart.Style, out Color? lineCol);
            return lineColor;
        }

        // From the docs: 'The default font color is the same as the Axis & Major Gridlines Line Color'
        internal Color? InbuiltStyleFontColor => GetDefaultFontColor();

        public ChartRenderer(ExcelChart chart, SvgRenderOptions options) : base(chart, options.FontTarget) 
        {
            SetChartArea(options);

            if (chart.HasTitle && chart.Series.Count > 0)
            {
                Title = new ChartTitleRenderer(this, (ExcelChartTitleStandard)chart.Title, "Chart Title");
            }
            else
            {
                Title = null;
            }

            //We need to create the plotarea before the legend and axes, as the trendlines can affect the value axis and should be rendererd in the legend.
            Plotarea = new ChartPlotareaRenderer(this)
            {
                ChartTypeDrawers = ChartTypeDrawer.Create(this)
            };

            if (chart.HasLegend)
            {
                Legend = new ChartLegendRenderer(this);
            }
            else
            {
                Legend = null;
            }

            if (Chart.Axis.Length != 0)
            {
                HorizontalAxis = GetAxis(false);
                VerticalAxis = GetAxis(true);
            }

            if (chart.Axis.Length > 2)
            {
                SecondVerticalAxis = GetAxis(true, 2);
                SecondHorizontalAxis = GetAxis(false, 2);
            }

            if (HasDataTable)
            {
                DataTable = new ChartDataTableRenderer(this);
            }

            Plotarea.SetPlotAreaRectangle();
            
            //As we need the plotarea dimensions to calculate the axis positions we need to set the axis positions after creating the plotarea.
            SetAxisPositionsFromPlotarea();

            Plotarea.DrawSeries();
            
            //Append all renderitems after everything has been created and positioned, to ensure the correct z-ordering.
            AppendItems();
        }
        private void SetAxisPositionsFromPlotarea()
        {
            if (VerticalAxis != null)
            {
                PlaceVerticalAxis(VerticalAxis, false);
            }

            if (HorizontalAxis != null)
            {
                PlaceHorizontalAxis(HorizontalAxis, false);
            }

            HorizontalAxis?.AddTickmarksAndValues(DefItems);
            VerticalAxis?.AddTickmarksAndValues(DefItems);

            if (SecondVerticalAxis != null)
            {
                PlaceVerticalAxis(SecondVerticalAxis, true);
                SecondVerticalAxis.AddTickmarksAndValues(DefItems);
            }   

            if (SecondHorizontalAxis != null)
            {
                PlaceHorizontalAxis(SecondHorizontalAxis, true);
                SecondHorizontalAxis.AddTickmarksAndValues(DefItems);
            }

            PlaceHorizontalAxisTitle(HorizontalAxis);
            PlaceHorizontalAxisTitle(SecondHorizontalAxis);
            PlaceVerticalAxisTitle(VerticalAxis);
            PlaceVerticalAxisTitle(SecondVerticalAxis);
        }
        private void PlaceHorizontalAxis(ChartAxisRenderer horizontalAxis, bool isSecondary)
        {
            var ax = horizontalAxis.Axis;
            var axisPos = ax.ActualAxisPosition;
            if (horizontalAxis.Axis.Deleted == false && horizontalAxis.Rectangle != null)
            {
                horizontalAxis.Rectangle.Width = Plotarea.Rectangle.Width;
                horizontalAxis.Rectangle.Left = Plotarea.Group.Left;
                horizontalAxis.Line?.X1 = (float)0;
                horizontalAxis.Line?.X2 = (float)horizontalAxis.Rectangle.Width;

                if (horizontalAxis.Line != null)
                {
                    if (ax.CrossesAt.HasValue)
                    {
                        double axisY;
                        if (isSecondary)
                        {
                            axisY = SecondVerticalAxis.GetPositionInPlotarea(ax.CrossesAt.Value);
                        }
                        else
                        {
                            axisY = VerticalAxis.GetPositionInPlotarea(ax.CrossesAt.Value);
                        }
                        horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = axisY;
                    }
                    else
                    {
                        if (axisPos == eAxisPosition.Bottom)
                        {
                            horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = Plotarea.Rectangle.Height;
                        }
                        else
                        {
                            horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = 0;
                            horizontalAxis.AdjustPosBySize = true;
                        }
                    }
                }

                switch (horizontalAxis.Axis.LabelPosition)
                {
                    case eTickLabelPosition.NextTo:
                        horizontalAxis.Rectangle.Top = Plotarea.Rectangle.GlobalTop + horizontalAxis.Line.Y1;
                        break;
                    case eTickLabelPosition.Low:
                        if (axisPos == eAxisPosition.Bottom)
                        {
                            horizontalAxis.Rectangle.Top = Plotarea.Rectangle.GlobalBottom;
                            horizontalAxis.AdjustPosBySize = false;
                        }
                        else
                        {
                            horizontalAxis.Rectangle.Top = Plotarea.Rectangle.GlobalTop;
                            horizontalAxis.AdjustPosBySize = true;
                        }
                        break;
                    case eTickLabelPosition.High:
                        if (axisPos == eAxisPosition.Bottom)
                        {
                            horizontalAxis.Rectangle.Top = Plotarea.Rectangle.GlobalTop;
                            horizontalAxis.AdjustPosBySize = true;
                        }
                        else
                        {
                            horizontalAxis.Rectangle.Top = Plotarea.Rectangle.GlobalBottom;
                            horizontalAxis.AdjustPosBySize = false;
                        }
                        break;
                }
            }
        }

        private void PlaceHorizontalAxisTitle(ChartAxisRenderer horizontalAxis)
        {
            if(horizontalAxis == null || horizontalAxis.Title == null)
            {
                return;
            }
            horizontalAxis.Title.Rectangle.Height = ChartArea.Rectangle.Height / 4;
            horizontalAxis.Title.Rectangle.Width = horizontalAxis.Rectangle?.Width ?? Plotarea.Rectangle.Width;

            if (horizontalAxis.Axis.AxisPosition == eAxisPosition.Top)
            {
                horizontalAxis.Title.TextBox.Top = GetBottomAxisTitlePos() - horizontalAxis.Title.TextBox.GetActualHeight() - 1.5;
            }
            else
            {
                horizontalAxis.Title.TextBox.Left = GetTopAxisTitlePos() - 1.5;
            }

            //if (horizontalAxis.Axis.Deleted)
            //{
            //    if (horizontalAxis.Axis.AxisPosition == eAxisPosition.Bottom)
            //    {
            //        if(horizontalAxis==SecondHorizontalAxis)
            //        {
            //            horizontalAxis.Title.TextBox.Top = Plotarea.Group.Top + Plotarea.Rectangle.Height + (HorizontalAxis?.Rectangle?.Height??0) + (HorizontalAxis?.Title?.Rectangle.Height ?? 0);
            //        }
            //        else
            //        {
            //            horizontalAxis.Title.TextBox.Top = Plotarea.Group.Top + Plotarea.Rectangle.Height;
            //        }
            //    }
            //    else
            //    {
            //        horizontalAxis.Title.TextBox.Top = Plotarea.Group.Top - horizontalAxis.Rectangle.Height;
            //    }
            //}
            //else
            //{
            //    if (horizontalAxis.Axis.AxisTextboxPosition == eActualAxisPosition.Bottom)
            //    {
            //        if (SecondHorizontalAxis != null && SecondHorizontalAxis.Axis.AxisTextboxPosition == eActualAxisPosition.Bottom)
            //        {
            //            horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Bottom + SecondHorizontalAxis.Rectangle.Height;
            //        }
            //        else if (horizontalAxis.Axis.AxisTextboxPosition == eActualAxisPosition.Bottom)
            //        {
            //            horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Bottom;
            //        }
            //        else
            //        {
            //            if (Legend != null && Chart.Legend.Position == eLegendPosition.Top)
            //            {
            //                horizontalAxis.Title.TextBox.Top = Legend.Rectangle.Top - horizontalAxis.Title.Rectangle.Height;
            //            }
            //            else
            //            {
            //                horizontalAxis.Title.TextBox.Top = ChartArea.Rectangle.Bottom - horizontalAxis.Title.Rectangle.Height - ChartArea.BottomMargin;
            //            }
            //        }
            //    }
            //    else
            //    {
            //        if(horizontalAxis==SecondHorizontalAxis && horizontalAxis.Axis.AxisTextboxPosition == eActualAxisPosition.Top && horizontalAxis.Axis.AxisTextboxPosition==eActualAxisPosition.Top)
            //        {
            //            horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Top - horizontalAxis.Title.Rectangle.Height;
            //        }
            //        else if (horizontalAxis.Axis.AxisTextboxPosition == eActualAxisPosition.Top)
            //        {
            //            horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Top - horizontalAxis.Title.TextBox.Height;
            //        }
            //        else
            //        {
            //            if (Legend != null && Chart.Legend.Position == eLegendPosition.Top)
            //            {
            //                horizontalAxis.Title.TextBox.Top = Legend.Rectangle.Bottom;
            //            }
            //            else if (Title != null)
            //            {
            //                horizontalAxis.Title.TextBox.Top = Title.Rectangle.Bottom;
            //            }
            //            else
            //            {
            //                horizontalAxis.Title.TextBox.Top = ChartArea.TopMargin;
            //            }
            //        }
            //    }
            //}
            horizontalAxis.Title.TextBox.Left = Plotarea.Group.Left + (Plotarea.Rectangle.Width / 2) - (horizontalAxis.Title.TextBox.Width / 2);
        }
        private void PlaceVerticalAxis(ChartAxisRenderer verticalAxis, bool isSecondary)
        {
            if (verticalAxis.Axis.Deleted == false && verticalAxis.Rectangle != null)
            {
                verticalAxis.Rectangle.Top = Plotarea.Group.Top;
                verticalAxis.Rectangle.Height = Plotarea.Rectangle.Height;
                verticalAxis.Line?.Y1 = (float)0;
                verticalAxis.Line?.Y2 = (float)verticalAxis.Rectangle.Height;
                var axisPos = verticalAxis.Axis.ActualAxisPosition;
                var ax = verticalAxis.Axis;

                if (ax.CrossesAt.HasValue)
                {
                    double axisX;
                    if (isSecondary)
                    {
                        axisX = SecondHorizontalAxis.GetPositionInPlotarea(ax.CrossesAt.Value);
                    }
                    else
                    {
                        axisX = HorizontalAxis.GetPositionInPlotarea(ax.CrossesAt.Value);
                    }

                    if(axisX<0)
                    {
                        axisX = 0;
                    }
                    else if(axisX > Plotarea.Rectangle.Width)
                    {
                        axisX = Plotarea.Rectangle.Width;
                    }
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = axisX;
                }
                else
                {
                    if (axisPos == eAxisPosition.Right)
                    {
                        verticalAxis.Line.X1 = verticalAxis.Line.X2 = Plotarea.Rectangle.Width;
                    }
                    else
                    {
                        verticalAxis.Line.X1 = verticalAxis.Line.X2 = 0;
                        verticalAxis.AdjustPosBySize = true;
                    }
                }
                switch(verticalAxis.Axis.LabelPosition)
                {
                    case eTickLabelPosition.NextTo:
                        verticalAxis.Rectangle.Left = Plotarea.Rectangle.GlobalLeft + verticalAxis.Line.X1;
                        break;
                    case eTickLabelPosition.Low:
                        if (axisPos == eAxisPosition.Left)
                        {
                            verticalAxis.Rectangle.Left = Plotarea.Rectangle.GlobalRight;
                            verticalAxis.AdjustPosBySize = false;
                        }
                        else
                        {
                            verticalAxis.Rectangle.Left = Plotarea.Rectangle.GlobalLeft;
                            verticalAxis.AdjustPosBySize = true;
                        }
                        break;
                    case eTickLabelPosition.High:
                        if (axisPos == eAxisPosition.Left)
                        {
                            verticalAxis.Rectangle.Left = Plotarea.Rectangle.GlobalRight;
                            verticalAxis.AdjustPosBySize = false;
                        }
                        else
                        {
                            verticalAxis.Rectangle.Left = Plotarea.Rectangle.GlobalLeft;
                            verticalAxis.AdjustPosBySize = true;
                        }
                        break;
                }
            }

            PlaceVerticalAxisTitle(verticalAxis);
        }

        private void PlaceVerticalAxisTitle(ChartAxisRenderer verticalAxis)
        {
            if (verticalAxis.Title != null)
            {
                var sinRot = Math.Abs(Math.Sin(MathHelper.Radians(verticalAxis.Title.TextBox.Rotation)));
                var cosRot = Math.Abs(Math.Cos(MathHelper.Radians(verticalAxis.Title.TextBox.Rotation)));
                verticalAxis.Title.TextBox.Top = Plotarea.Group.Top + (Plotarea.Rectangle.Height / 2) + ((verticalAxis.Title.TextBox.Height * cosRot + verticalAxis.Title.TextBox.Width * sinRot) / 2);

                if (verticalAxis.Axis.AxisPosition == eAxisPosition.Left)
                {
                    verticalAxis.Title.TextBox.Left = GetLeftAxisTitlePos() - verticalAxis.Title.TextBox.GetActualWidth() - 1.5;
                }
                else
                {
                    verticalAxis.Title.TextBox.Left = GetRightAxisTitlePos() - 1.5;
                }
            }
        }

        private double GetLeftAxisTitlePos()
        {
            if(VerticalAxis==null)
            {
                if(SecondVerticalAxis==null)
                {
                    return Plotarea.Rectangle.Left;
                }
                else
                {
                    return Math.Min(Plotarea.Rectangle.GlobalLeft, SecondVerticalAxis.Rectangle.Left);
                }
            }
            else
            {
                if (SecondVerticalAxis == null)
                {
                    return Math.Min(Plotarea.Rectangle.GlobalLeft, VerticalAxis.Rectangle.Left);
                }
                else
                {
                    return Math.Min(Plotarea.Rectangle.GlobalLeft, Math.Min(SecondVerticalAxis.Rectangle.Left, VerticalAxis.Rectangle.Left));
                }
            }
        }
        private double GetRightAxisTitlePos()
        {
            if (VerticalAxis == null)
            {
                if (SecondVerticalAxis == null)
                {
                    return Plotarea.Rectangle.GlobalRight;
                }
                else
                {
                    return Math.Max(Plotarea.Rectangle.GlobalRight, SecondVerticalAxis.Rectangle.Right);
                }
            }
            else
            {
                if (SecondVerticalAxis == null)
                {
                    return Math.Max(Plotarea.Rectangle.GlobalRight, VerticalAxis.Rectangle.Right);
                }
                else
                {
                    return Math.Max(Plotarea.Rectangle.GlobalRight, Math.Max(SecondVerticalAxis.Rectangle.Right, VerticalAxis.Rectangle.Right));
                }
            }
        }
        private double GetTopAxisTitlePos()
        {
            if (HorizontalAxis == null)
            {
                if (SecondHorizontalAxis == null)
                {
                    return Plotarea.Rectangle.GlobalTop;
                }
                else
                {
                    return Math.Min(Plotarea.Rectangle.GlobalTop, SecondHorizontalAxis.Rectangle.Top);
                }
            }
            else
            {
                if (SecondHorizontalAxis == null)
                {
                    return Math.Min(Plotarea.Rectangle.GlobalTop, HorizontalAxis.Rectangle.Top);
                }
                else
                {
                    return Math.Min(Plotarea.Rectangle.GlobalTop, Math.Min(SecondHorizontalAxis.Rectangle.Top, HorizontalAxis.Rectangle.Top));
                }
            }
        }
        private double GetBottomAxisTitlePos()
        {
            if (HorizontalAxis == null)
            {
                if (SecondHorizontalAxis == null)
                {
                    return Plotarea.Rectangle.Bottom;
                }
                else
                {
                    return Math.Max(Plotarea.Rectangle.GlobalBottom, SecondHorizontalAxis.Rectangle.Bottom);
                }
            }
            else
            {
                if (SecondHorizontalAxis == null)
                {
                    return Math.Max(Plotarea.Rectangle.GlobalBottom, HorizontalAxis.Rectangle.Bottom);
                }
                else
                {
                    return Math.Max(Plotarea.Rectangle.GlobalBottom, Math.Max(SecondHorizontalAxis.Rectangle.Bottom, VerticalAxis.Rectangle.Bottom));
                }
            }
        }

        private void SetChartArea(SvgRenderOptions options)
        {
            var item = new ChartAreaRenderer(this, options);
            item.Rectangle.Width = Bounds.Width;
            item.Rectangle.Height = Bounds.Height;            
            item.Rectangle.Style.SetDrawingPropertiesFill(Theme, Chart.Fill, Chart.StyleManager.Style?.ChartArea.FillReference.Color, UserSpaceSettings.UserSpaceOnUse_Global, item.GetDefaultFillColor());
            item.Rectangle.Name = "Plotarea";
            var borderstyle = Theme.FormatScheme.BorderStyle[0];

            //First check
            //1. Chart.Border (make sure to note the chart style ID
            //2. Chart.StyleManager.ChartArea.BorderReference
            //3. Theme.FormatScheme.BorderStyle[0] for subtle, [1] Moderate [2] Intense
            //4. If none of these contain even an empty node for the relevant property, Fallback to hardcoded documentation defaults 

            //Starting with: 1.1. noFill 1.2. solidFill (color = schemeColor (bg1)) 1.3. gradFill (gsList(black, white), shade(lin), flip = none, rotWithShape = true ) etc.

            //Note that a NoFill node for Charts means Transparent and that no nodes at all become bg1 as shown above

            var styleType = Chart.Style;

            var reference = Chart.StyleManager.Style?.ChartArea.BorderReference;

            //var chartStyleId = Chart.StyleManager.Style.Id;

            item.Rectangle.Style.SetDrawingPropertiesBorder(
                Theme,
                Chart.Border,
                reference?.Color,
                Chart.Border.Fill.Style != eFillStyle.NoFill,
                () => item.GetDefaultBorderColor(), 0.75d);

            item.Rectangle.RoundedCornerRadius = Chart.RoundedCorners ? 9 : 0;
            item.AppendRenderItems(RenderItems);
            item.SetMargins(Chart.TextBody);
            ChartArea = item;
        }

        private ChartAxisRenderer GetAxis(bool vertical, int offset = 0)
        {
            var axis = (ExcelChartAxisStandard)Chart.Axis[offset];
            if (axis.IsVertical == vertical)
            {
                return new ChartAxisRenderer(this, axis);
            }
            else if (Chart.Axis.Length > offset + 1)
            {
                axis = (ExcelChartAxisStandard)Chart.Axis[offset + 1];
                if (axis.IsVertical == vertical)
                {
                    return new ChartAxisRenderer(this, axis);
                }
            }
            return null;
        }

        public ExcelChart Chart 
        { 
            get =>(ExcelChart)Drawing;  
        }
        internal ChartDrawingObjectWithBackground ChartArea { get; set; }
        internal ChartLegendRenderer Legend { get; set; }
        internal ChartTitleRenderer Title { get; set; }
        internal ChartPlotareaRenderer Plotarea { get; set; }
        internal ChartAxisRenderer VerticalAxis { get; set; }
        internal ChartAxisRenderer HorizontalAxis { get; set; }
        internal ChartAxisRenderer SecondVerticalAxis { get; set; }
        internal ChartAxisRenderer SecondHorizontalAxis { get; set; }
        
        internal List<RenderItem> DefItems { get; } = new List<RenderItem>();
        public bool HasDataTable { get => Chart.PlotArea.DataTable != null;  }
        public ChartDataTableRenderer DataTable { get; private set; }

        internal void AddDefs(RenderItem item)
        {
            DefItems.Add(item);
        }
        public bool AppendItems()
        {
            Plotarea?.AppendRenderItems(RenderItems);

            HorizontalAxis?.AppendRenderItems(RenderItems);
            VerticalAxis?.AppendRenderItems(RenderItems);
            SecondHorizontalAxis?.AppendRenderItems(RenderItems);
            SecondVerticalAxis?.AppendRenderItems(RenderItems);

            if (Plotarea != null)
            {
                //TODO: Add typeDrawers group?
                //Currently drawer.AppendRenderItems rarely if ever actually adds anything
                //It goes directly on PlotArea mostly
                foreach (var drawer in Plotarea?.ChartTypeDrawers)
                {
                    drawer.AppendRenderItems(RenderItems);
                }
            }

            HorizontalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            VerticalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            SecondHorizontalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            SecondVerticalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            


            DataTable?.AppendRenderItems(RenderItems);
            Title?.AppendRenderItems(RenderItems);
            Legend?.AppendRenderItems(RenderItems);

            return true;
        }
        internal double GetPlotAreaTop()
        {
            var margin = 14D;
            if (Legend != null && Chart.Legend.Position == eLegendPosition.Top)
            {
                return Legend.Rectangle.Bottom + margin;
            }
            else if (Title != null)
            {
                return Title.Rectangle.Bottom + margin;
            }
            else
            {
                return margin;
            }

        }
    }
}
