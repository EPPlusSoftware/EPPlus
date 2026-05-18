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
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChart : ChartRenderer
    {
        public SvgChart(ExcelChart chart/*, IChartRenderer renderer*/) : base(chart)
        {
            SetChartArea();

            if(chart.HasTitle && chart.Series.Count > 0)
            {
                Title = new SvgChartTitle(this, (ExcelChartTitleStandard)chart.Title, "Chart Title");
            }
            else
            {
                Title = null;
            }

            //We need to create the plotarea before the legend and axes, as the trendlines can affect the value axis and should be rendererd in the legend.
            Plotarea = new SvgChartPlotarea(this);
            Plotarea.ChartTypeDrawers = ChartTypeDrawer.Create(this);

            if (chart.HasLegend)
            {
                Legend = new SvgChartLegend(this);
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

            if(chart.Axis.Length > 2)
            {
                SecondVerticalAxis = GetAxis(true, 2);
                SecondHorizontalAxis = GetAxis(false, 2);
            }

            Plotarea.SetPlotAreaRectangle(this);

            //As we need the plotarea dimensions to calculate the axis positions we need to set the axis positions after creating the plotarea.
            SetAxisPositionsFromPlotarea(this);
            
            Plotarea.DrawSeries();
        }

        private SvgChartAxis GetAxis(bool vertical, int offset=0)
        {
            var axis = (ExcelChartAxisStandard)Chart.Axis[offset];
            if(axis.IsVertical==vertical)
            {
                return new SvgChartAxis(this, axis);
            }
            else if(Chart.Axis.Length > offset + 1)
            {
                axis = (ExcelChartAxisStandard)Chart.Axis[offset + 1];
                if(axis.IsVertical==vertical)
                {
                    return new SvgChartAxis(this, axis);
                }
            }
            return null;
        }

        private void  SetAxisPositionsFromPlotarea(SvgChart sc)
        {
            if(Legend !=null && (sc.Chart.Legend.Position==eLegendPosition.Left || sc.Chart.Legend.Position == eLegendPosition.Right))
            {
                Legend.Rectangle.Top = sc.Plotarea.Rectangle.Top + (sc.Plotarea.Rectangle.Height / 2) - (Legend.Rectangle.Height / 2);
            }
            if(VerticalAxis != null)
            {
                PlaceVerticalAxis(sc, VerticalAxis);
                VerticalAxis.AddTickmarksAndValues(DefItems);
            }

            if (HorizontalAxis!=null && HorizontalAxis.Rectangle != null)
            {
                PlaceHorizontalAxis(sc, HorizontalAxis);

                //Make sure the horizontal axis is moved up if the vertical axis has a negative minimum value, so that the 0 value is at the correct position.
                if (VerticalAxis.Axis.AxisType == eAxisType.Val && VerticalAxis.Min < 0D)
                {
                    var newtop = VerticalAxis.GetPositionInPlotarea(0D) + sc.Plotarea.Rectangle.Top;
                    var topDiff = HorizontalAxis.Rectangle.Top - newtop;
                    HorizontalAxis.Rectangle.Top = newtop;
                    HorizontalAxis.Rectangle.Height += topDiff;
                    HorizontalAxis.Line.Y1 = HorizontalAxis.Line.Y2 = newtop;
                }

                HorizontalAxis.AddTickmarksAndValues(DefItems);
            }

            if (SecondVerticalAxis!=null)
            {
                PlaceVerticalAxis(sc, SecondVerticalAxis);
                SecondVerticalAxis.AddTickmarksAndValues(DefItems);
            }

            if (SecondHorizontalAxis != null && SecondHorizontalAxis.Rectangle != null)
            {
                PlaceHorizontalAxis(sc, SecondHorizontalAxis);
                SecondHorizontalAxis.AddTickmarksAndValues(DefItems);
            }
        }

        private void PlaceHorizontalAxis(SvgChart sc, SvgChartAxis horizontalAxis)
        {
            horizontalAxis.Rectangle.Width = Plotarea.Rectangle.Width;
            horizontalAxis.Rectangle.Left = Plotarea.Rectangle.Left;
            horizontalAxis.Line.X1 = (float)horizontalAxis.Rectangle.Left;
            horizontalAxis.Line.X2 = (float)horizontalAxis.Rectangle.Right;

            if (horizontalAxis.Axis.AxisPosition == eAxisPosition.Bottom)
            {
                horizontalAxis.Rectangle.Top = Plotarea.Rectangle.Bottom;
                horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = (float)Plotarea.Rectangle.Bottom;
            }
            else
            {
                horizontalAxis.Rectangle.Top = Plotarea.Rectangle.Top - horizontalAxis.Rectangle.Height;
                horizontalAxis.Line.Y1 = horizontalAxis.Line.Y2 = (float)Plotarea.Rectangle.Top;
            }

            if (horizontalAxis.Title != null)
            {
                horizontalAxis.Title.Rectangle.Height = sc.Bounds.Height / 4;
                horizontalAxis.Title.Rectangle.Width = horizontalAxis.Rectangle?.Width ?? sc.Plotarea.Rectangle.Width;
                //horizontalAxis.Title.InitTextBox();
                if (horizontalAxis.Axis.AxisPosition == eAxisPosition.Bottom)
                {
                    horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Bottom;
                }
                else
                {
                    horizontalAxis.Title.TextBox.Top = horizontalAxis.Rectangle.Top - horizontalAxis.Title.TextBox.Height;
                }
                horizontalAxis.Title.TextBox.Left = Plotarea.Rectangle.Left + (Plotarea.Rectangle.Width / 2) - (horizontalAxis.Title.TextBox.Width / 2);
            }
        }

        private void PlaceVerticalAxis(SvgChart sc, SvgChartAxis verticalAxis)
        {
            if (verticalAxis.Rectangle != null)
            {
                verticalAxis.Rectangle.Top = Plotarea.Rectangle.Top;
                verticalAxis.Rectangle.Height = Plotarea.Rectangle.Height;
                verticalAxis.Line.Y1 = (float)verticalAxis.Rectangle.Top;
                verticalAxis.Line.Y2 = (float)verticalAxis.Rectangle.Bottom;
                if (verticalAxis.Axis.AxisPosition == eAxisPosition.Left)
                {
                    verticalAxis.Rectangle.Left = Plotarea.Rectangle.Left - verticalAxis.Rectangle.Width;
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = (float)Plotarea.Rectangle.Left;
                }
                else
                {
                    verticalAxis.Rectangle.Left = Plotarea.Rectangle.Right;
                    verticalAxis.Line.X1 = verticalAxis.Line.X2 = (float)Plotarea.Rectangle.Right;
                }
            }

            if (verticalAxis.Title != null)
            {
                //verticalAxis.Title.Rectangle.Height = Plotarea.Rectangle.Height;
                //verticalAxis.Title.Rectangle.Width = sc.Bounds.Width / 4;
                //verticalAxis.Title.InitTextBox();
                var sinRot = Math.Abs(Math.Sin(MathHelper.Radians(verticalAxis.Title.TextBox.Rotation)));
                var cosRot = Math.Abs(Math.Cos(MathHelper.Radians(verticalAxis.Title.TextBox.Rotation)));
                verticalAxis.Title.TextBox.Top = Plotarea.Rectangle.Top + (Plotarea.Rectangle.Height / 2) + ((verticalAxis.Title.TextBox.Height * cosRot + verticalAxis.Title.TextBox.Width * sinRot) / 2);

                if (verticalAxis.Axis.AxisPosition == eAxisPosition.Left)
                {
                    if (verticalAxis.Rectangle == null)
                    {
                        verticalAxis.Title.TextBox.Left = sc.Plotarea.Rectangle.Left - verticalAxis.Title.TextBox.GetActualWidth() - 1.5;
                    }
                    else
                    {
                        verticalAxis.Title.TextBox.Left = verticalAxis.Rectangle.Left - verticalAxis.Title.TextBox.GetActualWidth() - 1.5;
                    }
                }
                else 
                {
                    if (verticalAxis.Rectangle == null)
                    {
                        verticalAxis.Title.TextBox.Left = Plotarea.Rectangle.Right;
                    }
                    else
                    {
                        verticalAxis.Title.TextBox.Left = verticalAxis.Rectangle.Right;
                    }
                }
                //verticalAxis.Title.RenderTextbox.TextAnchor = eTextAnchor.Middle;
            }
        }

        internal SvgChartObject ChartArea { get; set; }
        internal SvgChartLegend Legend { get; set; }
        internal SvgChartTitle Title { get; set; }
        internal SvgChartPlotarea Plotarea { get; set; }
        internal SvgChartAxis VerticalAxis { get; set; }
        internal SvgChartAxis HorizontalAxis { get; set; }
        internal SvgChartAxis SecondVerticalAxis { get; set; }
        internal SvgChartAxis SecondHorizontalAxis { get; set; }
        private void SetChartArea()
        {
            var item = new SvgChartArea(this);
            item.Rectangle.Width = Bounds.Width;
            item.Rectangle.Height = Bounds.Height;
            item.Rectangle.SetDrawingPropertiesFill(Chart.Fill, Chart.StyleManager.Style.ChartArea.FillReference.Color);
            item.Rectangle.SetDrawingPropertiesBorder(Chart.Border, Chart.StyleManager.Style.ChartArea.BorderReference.Color, Chart.Border.Width > 0);
            item.AppendRenderItems(RenderItems);
            ChartArea = item;
        }
        internal List<RenderItem> DefItems { get; } = new List<RenderItem>();
        internal void AddDefs(RenderItem item)
        {
            DefItems.Add(item);
        }
        public void Render(StringBuilder sb)
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

            sb.Append($"<svg width=\"{Bounds.Width.PointToPixelString()}\" height=\"{Bounds.Height.PointToPixelString()}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" Overflow=\"Hidden\" >");
            //Write defs used for gradient colors
            var writer = new SvgDrawingWriter(this);
            writer.WriteSvgDefs(sb, RenderItems);

            foreach (var item in RenderItems)
            {
                item.Render(sb);
            }

            sb.Append("</svg>");
        }

        internal double GetPlotAreaTop()
        {
            var margin = 14D;
            if(Legend!=null && Chart.Legend.Position==eLegendPosition.Top)
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

        internal SvgRenderLineItem GetSeriesIcon(ExcelChartStandardSerie s, int index, BoundingBox parentItem)
        {
            const float MarginExtra = 1.5f;
            const float LineLength = 21;

            var item = new SvgRenderLineItem(this, parentItem);
            item.SetDrawingPropertiesFill(s.Fill, this.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
            item.SetDrawingPropertiesBorder(s.Border, this.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, s.Border.Fill.Style != eFillStyle.NoFill, 0.75);

            float y = (float)parentItem.Top + MarginExtra;
            float x = 0;
            item.X1 = x;
            item.Y1 = y;
            item.X2 = x + LineLength;
            item.Y2 = y;
            item.LineCap = eLineCap.Round;

            return item;
        }
    }

    internal interface IChartRenderer
    {
        DrawingObject ChartAreaRenderer { get; }
        DrawingObject TitleRenderer { get; }
        DrawingObject LegendRenderer { get; }
        DrawingObject AxisRenderer { get; }
        DrawingObject AxisTextboxRenderer { get; }
        DrawingObject PlotareaRenderer { get; }
    }
}