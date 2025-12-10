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

            if(chart.YAxis.Deleted==false)
            {
                VerticalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.YAxis);
            }

            if (chart.YAxis.Deleted == false)
            {
                HorizontalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.XAxis);
            }
            
            if(chart.Axis.Length > 2 && chart.Axis[2].Deleted==false)
            {
                SecondHorizontalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.Axis[2]);
            }
            
            Plotarea = new SvgChartPlotarea(this);            
        }

        internal SvgRenderRectItem ChartArea { get; set; }
        internal SvgChartLegend Legend { get; set; }
        internal SvgChartTitle Title { get; set; }
        internal SvgChartPlotarea Plotarea { get; set; }
        internal SvgChartAxis VerticalAxis { get; set; }
        internal SvgChartAxis HorizontalAxis { get; set; }
        internal SvgChartAxis SecondHorizontalAxis { get; set; }
        internal SvgChartTitle VerticalAxisTitle { get; set; }
        internal SvgChartTitle HorizontalAxisTitle { get; set; }
        internal SvgChartTitle SecondHorizontalAxisTitle { get; set; }
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
            sb.Append($"<svg width=\"{Size.Width}\" height=\"{Size.Height}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"preserve\" overflow=\"hidden\" >");
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
            
            Title?.Render(sb);
            Legend?.Render(sb);

            if (gItemTest != null)
            {
                gItemTest.RenderEndGroup(sb);
            }

            HorizontalAxis?.Render(sb);
            HorizontalAxisTitle?.Render(sb);
            VerticalAxis?.Render(sb);
            VerticalAxisTitle?.Render(sb);
            SecondHorizontalAxis?.Render(sb);
            SecondHorizontalAxisTitle?.Render(sb);

            sb.Append("</svg>");
        }

        internal double GetPlotAreaTop()
        {
            var topMargin = 14;
            if (Title == null)
            {
                return topMargin; //
            }
            else
            {
                return Title.Rectangle.Bottom + topMargin;
            }

        }
        internal double GetPlotAreaBottom()
        {
            return 0;
        }
    }
}