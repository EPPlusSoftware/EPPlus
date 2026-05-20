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
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using System.Text;
using d=OfficeOpenXml.Drawing.Renderer;
namespace EPPlusImageRenderer
{
    internal class ChartRenderer : d.DrawingRenderer
    {       
        public ChartRenderer(ExcelChart chart) : base(chart) 
        {
            SetChartArea();

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

            Plotarea.SetPlotAreaRectangle(this);

            //As we need the plotarea dimensions to calculate the axis positions we need to set the axis positions after creating the plotarea.
            SetAxisPositionsFromPlotarea(this);

            Plotarea.DrawSeries();
        }
        private void SetChartArea()
        {
            var item = new ChartAreaRenderer(this);
            item.Rectangle.Width = Bounds.Width;
            item.Rectangle.Height = Bounds.Height;
            item.Rectangle.SetDrawingPropertiesFill(Theme, Chart.Fill, Chart.StyleManager.Style.ChartArea.FillReference.Color);
            item.Rectangle.SetDrawingPropertiesBorder(Theme, Chart.Border, Chart.StyleManager.Style.ChartArea.BorderReference.Color, Chart.Border.Width > 0);
            item.AppendRenderItems(RenderItems);
            ChartArea = item;
        }

        public ExcelChart Chart 
        { 
            get =>(ExcelChart)Drawing;  
        }
        internal ChartDrawingObject ChartArea { get; set; }
        internal ChartLegendRenderer Legend { get; set; }
        internal ChartTitleRenderer Title { get; set; }
        internal ChartPlotareaRenderer Plotarea { get; set; }
        internal SvgChartAxis VerticalAxis { get; set; }
        internal SvgChartAxis HorizontalAxis { get; set; }
        internal SvgChartAxis SecondVerticalAxis { get; set; }
        internal SvgChartAxis SecondHorizontalAxis { get; set; }

        public string Render()
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
