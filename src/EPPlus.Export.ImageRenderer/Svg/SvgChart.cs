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
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChart : DrawingChart
    {
        public SvgChart(ExcelChart chart) : base(chart)
        {
            Chart = chart;

            if(chart.HasTitle)
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
            if(chart.Axis.Length > 2)
            {
                SecondHorizontalAxis = new SvgChartAxis(this, (ExcelChartAxisStandard)chart.Axis[2]);
            }
            //Plotarea = new SvgChartPlotarea(this);

            AddChartArea();
            AddPlotArea();  
            AddChartTitle();
            AddLegend();    
        }

        private void AddLegend()
        {
        }

        private void AddChartTitle()
        {
            RenderItems.Add(Title.Rectangle);
            RenderItems.Add(Title.TextBox);
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
        private void AddPlotArea()
        {
            if (HorizontalAxis!=null) RenderItems.AddRange(HorizontalAxis?.RenderItems);
            if (HorizontalAxisTitle != null) RenderItems.AddRange(HorizontalAxisTitle?.RenderItems);
            if (VerticalAxis != null) RenderItems.AddRange(VerticalAxis?.RenderItems);
            if (VerticalAxisTitle != null) RenderItems.AddRange(VerticalAxisTitle?.RenderItems);
            if (SecondHorizontalAxis != null) RenderItems.AddRange(SecondHorizontalAxis?.RenderItems);
            if (SecondHorizontalAxisTitle != null) RenderItems.AddRange(SecondHorizontalAxisTitle?.RenderItems);
        }
        private void AddChartArea()
        {
            var item = new SvgRenderRectItem(Chart);
            item.Width = Size.Width;
            item.Height = Size.Height;
            item.SetDrawingPropertiesFill(Chart.Fill, Chart.StyleManager.Style.ChartArea.FillReference.Color);
            item.SetDrawingPropertiesBorder(Chart.Border, Chart.StyleManager.Style.ChartArea.BorderReference.Color, Chart.Border.Width > 0);
            RenderItems.Add(item);
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
            
            if (gItemTest != null)
            {
                gItemTest.RenderEndGroup(sb);
            }
            sb.Append("</svg>");
        }
    }
}