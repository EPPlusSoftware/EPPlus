using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class BarColumnChartTypeDrawer : ChartTypeDrawer
    {
        List<SvgChartSerieDataLabel> serieDataLabels = new List<SvgChartSerieDataLabel>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();

        internal BarColumnChartTypeDrawer(SvgChart svgChart, ExcelBarChart chartType) : base(svgChart, chartType)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
            RenderItems.Add(groupItem);
            var xValues = new List<List<object>>();
            var yValues = new List<List<object>>();
            int serCounter = 0;

            foreach (ExcelBarChartSerie serie in chartType.Series)
            {
                var yValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                var xValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                xValues.Add(xValue);
                yValues.Add(yValue);

                //if (serie.HasDataLabel)
                //{
                //    var datalabel = new SvgChartSerieDataLabel(svgChart, serie.DataLabel, svgChart.Bounds, serie, xValue, yValue, serCounter);
                //    serieDataLabels.Add(datalabel);
                //}
                serCounter++;
            }

            if(chartType.IsTypeStacked())
            {
                SumSeries(yValues);
            }
            else if (chartType.IsTypePercentStacked())
            {
                ExcelChartAxisStandard.CalculateStacked100(yValues);
            }
            var count = Math.Min(xValues.Count, yValues.Count);
            for (var i= 0; i < xValues.Count; i++)  
            {
                var serie = (ExcelBarChartSerie)chartType.Series[i];

                var dataPoints = new List<BoundingBox>();

                AddColumn(chartType, serie, xValues, yValues, dataPoints, count, i);

                dataPointsPerSerie.Add(dataPoints);

                //if (serie.HasDataLabel)
                //{
                //    for (int j = 0; j < dataPoints.Count; j++)
                //    {
                //        serieDataLabels[i].SetParentPoint(dataPoints[j], j);
                //    }
                //}
            }
            RenderItems.Add(new SvgEndGroupItem(ChartRenderer, null));

            foreach (var dataLabel in serieDataLabels)
            {
                dataLabel.AppendRenderItems(RenderItems);
            }
        }
        private void AddColumn(ExcelBarChart chartType, ExcelBarChartSerie serie, List<List<object>> xSeries, List<List<object>> ySeries, List<BoundingBox> dataPoints, int seriesCount, int position)
        {            
            SvgChartAxis yAxis, xAxis;
            if (chartType.UseSecondaryAxis)
            {
                yAxis = _svgChart.SecondVerticalAxis;
                xAxis = _svgChart.SecondHorizontalAxis;
                if(xAxis.Axis.Deleted && xAxis.Values==null)
                {
                    xAxis = _svgChart.HorizontalAxis;
                }
            }
            else
            {
                yAxis = _svgChart.VerticalAxis;
                xAxis = _svgChart.HorizontalAxis;
            }
            var xValues = xSeries[position];
            var yValues = ySeries[position];
            
            var slotSize = yValues.Count;
            var gapPercent = chartType.GapWidth / 100D;     //Gap width between bars/columns in percent
            var overlapPercent = chartType.Overlap / 100D;  //Overlap  between bars/columns in percent
            var slotWidth = _svgChart.Plotarea.Rectangle.Width / slotSize;
            var clusterWidth = slotWidth * 100 / (100 + chartType.GapWidth);
            var step = 1 - overlapPercent;
            double barWidth;
            barWidth = slotWidth / (1 + (seriesCount - 1) * step + gapPercent);
            var halfGap = (barWidth * gapPercent) / 2;

            double yAxisStart;

            if (yAxis.Axis.Crosses == eCrosses.AutoZero)
            {
                yAxisStart = yAxis.GetPositionInPlotarea(yAxis.Min <= 0 ? 0D : yAxis.Min, true);
            }
            else if (yAxis.Axis.Crosses == eCrosses.Min)
            {
                yAxisStart = yAxis.GetPositionInPlotarea(yAxis.Min, true);
            }
            else
            {
                yAxisStart = yAxis.GetPositionInPlotarea(yAxis.Max, true);
            }

            var isStacked = chartType.IsTypeStacked();
            var isStacked100 = chartType.IsTypePercentStacked();
            for (var i = 0; i < yValues.Count; i++)
            {
                double x;
                if (xValues == null || xAxis.Axis.AxisType == eAxisType.Cat)
                {
                    x = (double)i;
                }
                else
                {
                    x = ConvertUtil.GetValueDouble(xValues[i], false, true);
                }

                var y = ConvertUtil.GetValueDouble(yValues[i], false, true);

                var xPos = xAxis.GetPositionInPlotarea(x, true) + halfGap + position * barWidth * step;
                var yPos = yAxis.GetPositionInPlotarea(y);

                var rect = new SvgRenderRectItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
                rect.Left = xPos;
                rect.Width = barWidth;
                if(position > 0 && (isStacked || isStacked100))
                {
                    //Stacked
                    var pYValues = ySeries[position - 1];
                    var pAxisEnd = ConvertUtil.GetValueDouble(pYValues[i]);
                    var yPrevPos = yAxis.GetPositionInPlotarea(pAxisEnd, false);
                    if (y < 0)
                    {
                        rect.Top = yPrevPos;
                        rect.Height = yPos - yPrevPos;
                    }
                    else
                    {
                        rect.Top = yPos;
                        rect.Height = yPrevPos - yPos;
                    }
                }
                else
                {
                    if (y < 0)
                    {
                        rect.Top = yAxisStart;
                        rect.Height = yPos - yAxisStart;
                    }
                    else
                    {
                        rect.Top = yPos;
                        rect.Height = yAxisStart - yPos;
                    }
                }
                rect.SetDrawingPropertiesFill(serie.Fill, chartType.StyleManager.Style.SeriesAxis.FillReference.Color);
                rect.SetDrawingPropertiesBorder(serie.Border, chartType.StyleManager.Style.SeriesAxis.BorderReference.Color, true);
                rect.SetDrawingPropertiesEffects(serie.Effect);
                RenderItems.Add(rect);
            }

        }
    }

}
