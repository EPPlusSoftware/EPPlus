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
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Text;
using d=OfficeOpenXml.Drawing.Renderer;
namespace EPPlusImageRenderer
{
    internal class ChartRenderer : d.DrawingRenderer
    {
        public ChartRenderer(ExcelChart chart, SvgRenderOptions options) : base(chart) 
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
            Plotarea = new ChartPlotareaRenderer(this);
            Plotarea.ChartTypeDrawers = ChartTypeDrawer.Create(this);

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
                PlaceVerticalAxis(VerticalAxis);
                //Make sure the horizontal axis is moved up if the vertical axis has a negative minimum value, so that the 0 value is at the correct position.
                if (VerticalAxis.Axis.TickLabelPosition == eTickLabelPosition.NextTo && HorizontalAxis.Axis.AxisType == eAxisType.Val && HorizontalAxis.Min < 0D)
                {
                    Plotarea.Rectangle.Width += VerticalAxis.Rectangle.Width;
                    Plotarea.Group.Left = VerticalAxis.Rectangle.Left;
                    var newRight = Plotarea.Group.Left + HorizontalAxis.GetPositionInPlotarea(0D);
                    var rightDiff = newRight - VerticalAxis.Rectangle.Width;
                    VerticalAxis.Rectangle.Left = rightDiff;
                    VerticalAxis.Line.X1 = VerticalAxis.Line.X2 = newRight;
                }
                VerticalAxis.AddTickmarksAndValues(DefItems);
            }

            if (HorizontalAxis != null && HorizontalAxis.Rectangle != null)
            {
                PlaceHorizontalAxis(HorizontalAxis, false);

                //Make sure the horizontal axis is moved up if the vertical axis has a negative minimum value, so that the 0 value is at the correct position.
                if (HorizontalAxis.Axis.TickLabelPosition == eTickLabelPosition.NextTo && VerticalAxis.Axis.AxisType == eAxisType.Val && VerticalAxis.Min < 0D)
                {
                    var newtop = VerticalAxis.GetPositionInPlotarea(0D) + Plotarea.Group.Top;
                    var topDiff = HorizontalAxis.Rectangle.Top - newtop;
                    HorizontalAxis.Rectangle.Top = newtop;
                    HorizontalAxis.Rectangle.Height += topDiff;
                    HorizontalAxis.Line.Y1 = HorizontalAxis.Line.Y2 = newtop;
                }

                HorizontalAxis.AddTickmarksAndValues(DefItems);
            }

            if (SecondVerticalAxis != null)
            {
                PlaceVerticalAxis(SecondVerticalAxis);
                SecondVerticalAxis.AddTickmarksAndValues(DefItems);
            }

            if (SecondHorizontalAxis != null && SecondHorizontalAxis.Rectangle != null)
            {
                PlaceHorizontalAxis(SecondHorizontalAxis, true);
                SecondHorizontalAxis.AddTickmarksAndValues(DefItems);
            }
        }

        private void PlaceHorizontalAxis(ChartAxisRenderer horizontalAxis, bool isSecondary)
        {
            if (horizontalAxis.Axis.Deleted == false)
            {
                horizontalAxis.Rectangle.Width = Plotarea.Rectangle.Width;
                horizontalAxis.Rectangle.Left = Plotarea.Group.Left;
                horizontalAxis.Line.X1 = (float)horizontalAxis.Rectangle.Left;
                horizontalAxis.Line.X2 = (float)horizontalAxis.Rectangle.Right;
                
                var axisPos = horizontalAxis.Axis.ActualAxisPosition;
                if (axisPos == eActualAxisPosition.Bottom)
                {
                    horizontalAxis.Rectangle.Top = Plotarea.Group.Top + Plotarea.Rectangle.Height;
                    horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = (float)Plotarea.Group.Top + Plotarea.Rectangle.Height;
                }
                else if(axisPos == eActualAxisPosition.BottomSecond)
                {
                    horizontalAxis.Rectangle.Top = Plotarea.Group.Top + Plotarea.Rectangle.Height + HorizontalAxis.Rectangle.Height;
                    horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = horizontalAxis.Rectangle.Top;
                }
                else if(axisPos == eActualAxisPosition.Top)
                {
                    horizontalAxis.Rectangle.Top = Plotarea.Group.Top - horizontalAxis.Rectangle.Height;
                    horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = (float)Plotarea.Group.Top;
                }
                else
                {
                    horizontalAxis.Rectangle.Top = Plotarea.Group.Top - horizontalAxis.Rectangle.Height - HorizontalAxis.Rectangle.Height;
                    horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = horizontalAxis.Rectangle.Bottom;
                }
            }
            if (horizontalAxis.Title != null)
            {
                PlaceHorizontalAxisTitle(horizontalAxis);
            }
        }

        private void PlaceHorizontalAxisTitle(ChartAxisRenderer horizontalAxis)
        {
            horizontalAxis.Title.Rectangle.Height = Bounds.Height / 4;
            horizontalAxis.Title.Rectangle.Width = horizontalAxis.Rectangle?.Width ?? Plotarea.Rectangle.Width;
            if (horizontalAxis.Axis.Deleted)
            {
                if (horizontalAxis.Axis.AxisPosition == eAxisPosition.Bottom)
                {
                    horizontalAxis.Title.TextBox.Top = Plotarea.Group.Top + Plotarea.Rectangle.Height;
                }
                else
                {
                    horizontalAxis.Title.TextBox.Top = Plotarea.Group.Top - horizontalAxis.Rectangle.Height;
                }
            }
            else
            {
                if (horizontalAxis.Axis.AxisPosition == eAxisPosition.Bottom)
                {
                    if (SecondHorizontalAxis != null && SecondHorizontalAxis.Axis.ActualAxisPosition == eActualAxisPosition.BottomSecond)
                    {
                        horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Bottom + SecondHorizontalAxis.Rectangle.Height;
                    }
                    else if (horizontalAxis.Axis.ActualAxisPosition == eActualAxisPosition.Bottom)
                    {
                        horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Bottom;
                    }
                    else
                    {
                        if (Legend != null && Chart.Legend.Position == eLegendPosition.Top)
                        {
                            horizontalAxis.Title.TextBox.Top = Legend.Rectangle.Top - horizontalAxis.Title.Rectangle.Height;
                        }
                        else
                        {
                            horizontalAxis.Title.TextBox.Top = ChartArea.Rectangle.Bottom - horizontalAxis.Title.Rectangle.Height - ChartArea.BottomMargin;
                        }
                    }
                }
                else
                {
                    if (SecondHorizontalAxis.Axis.ActualAxisPosition == eActualAxisPosition.TopSecond)
                    {
                        horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Top - SecondHorizontalAxis.Rectangle.Height - horizontalAxis.Title.TextBox.Height;
                    }
                    else if (horizontalAxis.Axis.ActualAxisPosition == eActualAxisPosition.Top)
                    {
                        horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Top - horizontalAxis.Title.TextBox.Height;
                    }
                    else
                    {
                        if (Legend != null && Chart.Legend.Position == eLegendPosition.Top)
                        {
                            horizontalAxis.Title.TextBox.Top = Legend.Rectangle.Bottom;
                        }
                        else if (Title != null)
                        {
                            horizontalAxis.Title.TextBox.Top = Title.Rectangle.Bottom;
                        }
                        else
                        {
                            horizontalAxis.Title.TextBox.Top = ChartArea.TopMargin;
                        }
                    }
                }
            }
            horizontalAxis.Title.TextBox.Left = Plotarea.Group.Left + (Plotarea.Rectangle.Width / 2) - (horizontalAxis.Title.TextBox.Width / 2);
        }

        private void PlaceVerticalAxis(ChartAxisRenderer verticalAxis)
        {
            if (verticalAxis.Axis.Deleted == false && verticalAxis.Rectangle != null)
            {
                verticalAxis.Rectangle.Top = Plotarea.Group.Top;
                verticalAxis.Rectangle.Height = Plotarea.Rectangle.Height;
                verticalAxis.Line.Y1 = (float)verticalAxis.Rectangle.Top;
                verticalAxis.Line.Y2 = (float)verticalAxis.Rectangle.Bottom;
                var axisPos = verticalAxis.Axis.ActualAxisPosition;

                if (axisPos == eActualAxisPosition.Left)
                {
                    verticalAxis.Rectangle.Left = Plotarea.Group.Left - verticalAxis.Rectangle.Width;
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = (float)Plotarea.Group.Left;
                }
                else if (axisPos == eActualAxisPosition.LeftSecond)
                {
                    verticalAxis.Rectangle.Left = Plotarea.Group.Left - verticalAxis.Rectangle.Width - VerticalAxis.Rectangle.Width;
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = (float)Plotarea.Group.Left;
                }
                else if (axisPos == eActualAxisPosition.Right)
                {
                    verticalAxis.Rectangle.Left = Plotarea.Group.Left + Plotarea.Rectangle.Width;
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = (float)Plotarea.Group.Left + Plotarea.Rectangle.Width;
                }
                else
                {
                    verticalAxis.Rectangle.Left = Plotarea.Group.Left + Plotarea.Rectangle.Width + VerticalAxis.Rectangle.Width;
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = (float)Plotarea.Group.Left + Plotarea.Rectangle.Width;
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
                    if (verticalAxis.Rectangle == null)
                    {
                        verticalAxis.Title.TextBox.Left = Plotarea.Group.Left - verticalAxis.Title.TextBox.GetActualWidth() - 1.5;
                    }
                    else
                    {
                        verticalAxis.Title.TextBox.Left = Plotarea.Group.Left - verticalAxis.Rectangle.Width - verticalAxis.Title.TextBox.GetActualWidth() - 1.5;
                    }
                }
                else
                {
                    if (verticalAxis.Rectangle == null)
                    {
                        verticalAxis.Title.TextBox.Left = Plotarea.Group.Left + Plotarea.Rectangle.Width;
                    }
                    else
                    {
                        verticalAxis.Title.TextBox.Left = verticalAxis.Rectangle.Right;
                    }
                }
            }
        }

        private void SetChartArea(SvgRenderOptions options)
        {
            var item = new ChartAreaRenderer(this, options);
            item.Rectangle.Width = Bounds.Width;
            item.Rectangle.Height = Bounds.Height;
            
            item.Rectangle.SetDrawingPropertiesFill(Theme, Chart.Fill, Chart.StyleManager.Style?.ChartArea.FillReference.Color, true, Theme.ColorScheme.Light1.GetColor());
            item.Rectangle.SetDrawingPropertiesBorder(Theme, Chart.Border, Chart.StyleManager.Style?.ChartArea.BorderReference.Color, Chart.Border.Width > 0);
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
        internal ChartDrawingObject ChartArea { get; set; }
        internal ChartLegendRenderer Legend { get; set; }
        internal ChartTitleRenderer Title { get; set; }
        internal ChartPlotareaRenderer Plotarea { get; set; }
        internal ChartAxisRenderer VerticalAxis { get; set; }
        internal ChartAxisRenderer HorizontalAxis { get; set; }
        internal ChartAxisRenderer SecondVerticalAxis { get; set; }
        internal ChartAxisRenderer SecondHorizontalAxis { get; set; }

        internal List<RenderItem> DefItems { get; } = new List<RenderItem>();
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
                foreach (var drawer in Plotarea?.ChartTypeDrawers)
                {
                    drawer.AppendRenderItems(RenderItems);
                }
            }

            HorizontalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            VerticalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            SecondHorizontalAxis?.Textboxes?.AppendRenderItems(RenderItems);
            SecondVerticalAxis?.Textboxes?.AppendRenderItems(RenderItems);

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
        internal LineRenderItem GetSeriesIcon(ExcelChartStandardSerie s, int index, BoundingBox parentItem)
        {
            const float MarginExtra = 1.5f;
            const float LineLength = 21;

            var item = new LineRenderItem(parentItem);
            item.SetDrawingPropertiesFill(Theme, s.Fill, Chart.StyleManager.Style.SeriesLine.FillReference.Color, false);
            item.SetDrawingPropertiesBorder(Theme, s.Border, Chart.StyleManager.Style.SeriesLine.BorderReference.Color, s.Border.Fill.Style != eFillStyle.NoFill, null, 0.75, false);

            float y = (float)parentItem.Top + MarginExtra;
            float x = 0;
            item.X1 = x;
            item.Y1 = y;
            item.X2 = x + LineLength;
            item.Y2 = y;
            item.LineCap = LineCap.Round;

            return item;
        }

    }
}
