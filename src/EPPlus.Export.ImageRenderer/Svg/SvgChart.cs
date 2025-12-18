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
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChart : DrawingChart
    {
        public SvgChart(ExcelChart chart) : base(chart)
        {
            Chart = chart;
            SetChartArea();

            if(chart.HasTitle && chart.Series.Count > 0)
            {
                Title = new SvgChartTitle(this, (ExcelChartTitleStandard)chart.Title, "Chart Title");
            }
            else
            {
                Title = null;
            }

            if (chart.HasLegend)
            {
                Legend = new SvgChartLegend(this);
            }
            else
            {
                Legend = null;
            }

            VerticalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.YAxis);
            HorizontalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.XAxis);
            if (chart.Axis.Length > 2)
            {
                SecondVerticalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.Axis[2]);
            }

            Plotarea = new SvgChartPlotarea(this);

            SetAxisPositionsFromPlotarea(this);
        }

        private void SetAxisPositionsFromPlotarea(SvgChart sc)
        {
            if(VerticalAxis != null)
            {
                if (VerticalAxis.Rectangle != null)
                {
                    VerticalAxis.Rectangle.Top = Plotarea.Rectangle.Top;
                    VerticalAxis.Rectangle.Height = Plotarea.Rectangle.Height;
                    VerticalAxis.Rectangle.Left = Plotarea.Rectangle.Left - VerticalAxis.Rectangle.Width;
                    VerticalAxis.Line.X1 = VerticalAxis.Line.X2 = (float)Plotarea.Rectangle.Left;
                    VerticalAxis.Line.Y1 = (float)VerticalAxis.Rectangle.Top;
                    VerticalAxis.Line.Y2 = (float)VerticalAxis.Rectangle.Bottom;
                }

                if(VerticalAxis.Title!=null)
                {
                    VerticalAxis.Title.Rectangle.Top = Plotarea.Rectangle.Top + (Plotarea.Rectangle.Height / 2) - (VerticalAxis.Title.Rectangle.Height / 2);
                    VerticalAxis.Title.InitTextBox();
                }
            }

            if (HorizontalAxis!=null)
            {
                HorizontalAxis.Rectangle.Top = Plotarea.Rectangle.Bottom;
                HorizontalAxis.Rectangle.Width = Plotarea.Rectangle.Width;
                HorizontalAxis.Rectangle.Left = Plotarea.Rectangle.Left;
                HorizontalAxis.Line.Y1 = HorizontalAxis.Line.Y2 = (float)Plotarea.Rectangle.Bottom;
                HorizontalAxis.Line.X1 = (float)HorizontalAxis.Rectangle.Left;
                HorizontalAxis.Line.X2 = (float)HorizontalAxis.Rectangle.Right;

                if (HorizontalAxis.Title != null)
                {
                    HorizontalAxis.Title.Rectangle.Left = Plotarea.Rectangle.Left + (Plotarea.Rectangle.Width / 2) - (VerticalAxis.Title.Rectangle.Width / 2);
                    HorizontalAxis.Title.InitTextBox();
                }
            }

            if (SecondVerticalAxis!=null)
            {
                SecondVerticalAxis.Rectangle.Top = Plotarea.Rectangle.Top;
                SecondVerticalAxis.Rectangle.Height = Plotarea.Rectangle.Height;
                SecondVerticalAxis.Rectangle.Left = Plotarea.Rectangle.Right;
                VerticalAxis.Line.X1 = VerticalAxis.Line.X2 = (float)Plotarea.Rectangle.Right;
                SecondVerticalAxis.Line.Y1 = (float)SecondVerticalAxis.Rectangle.Top;
                SecondVerticalAxis.Line.Y2 = (float)SecondVerticalAxis.Rectangle.Bottom;
                if (SecondVerticalAxis.Title != null)
                {
                    SecondVerticalAxis.Title.Rectangle.Top = Plotarea.Rectangle.Top + (Plotarea.Rectangle.Height / 2) - (VerticalAxis.Title.Rectangle.Height / 2);
                    SecondVerticalAxis.Title.InitTextBox();
                }
            }
        }

        internal SvgRenderRectItem ChartArea { get; set; }
        internal SvgChartLegend Legend { get; set; }
        internal SvgChartTitle Title { get; set; }
        internal SvgChartPlotarea Plotarea { get; set; }
        internal SvgChartAxis VerticalAxis { get; set; }
        internal SvgChartAxis HorizontalAxis { get; set; }
        internal SvgChartAxis SecondVerticalAxis { get; set; }
        internal SvgChartTitle VerticalAxisTitle { get; set; }
        internal SvgChartTitle HorizontalAxisTitle { get; set; }
        internal SvgChartTitle SecondVerticalAxisTitle { get; set; }
        private void SetChartArea()
        {
            var item = new SvgRenderRectItem(Chart);
            item.Width = Size.Width;
            item.Height = Size.Height;
            item.SetDrawingPropertiesFill(Chart.Fill, Chart.StyleManager.Style.ChartArea.FillReference.Color);
            item.SetDrawingPropertiesBorder(Chart.Border, Chart.StyleManager.Style.ChartArea.BorderReference.Color, Chart.Border.Width > 0);
            RenderItems.Add(item);
            ChartArea = item;
        }

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<svg width=\"{Size.Width}\" height=\"{Size.Height}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" Overflow=\"Hidden\" >");
            //Write defs used for gradient colors
            var writer = new SvgDrawingWriter(this);
            writer.WriteSvgDefs(sb, RenderItems);

            SvgGroupItem gItemTest = null;
            foreach (var item in RenderItems)
            {
                item.Render(sb);
                if (item.Type == SvgItemType.Group && gItemTest == null)
                {
                    gItemTest = (SvgGroupItem)item;
                }
            }
            
            if (gItemTest != null)
            {
                gItemTest.RenderEndGroup(sb);
            }

            Plotarea?.Render(sb);

            HorizontalAxis?.Render(sb);
            VerticalAxis?.Render(sb);
            SecondVerticalAxis?.Render(sb);

            Legend?.Render(sb);
            Title?.Render(sb);

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
    }
}