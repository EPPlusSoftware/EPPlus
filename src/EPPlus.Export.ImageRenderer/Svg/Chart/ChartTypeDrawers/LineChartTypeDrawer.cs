using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class LineChartTypeDrawer : ChartTypeDrawer
    {
        List<SvgChartSerieDataLabel> serieDataLabels = new List<SvgChartSerieDataLabel>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();
        internal override bool SupportsTrendlines => true;
        internal List<SvgTrendline> Trendlines { get; } = new List<SvgTrendline>();
        internal LineChartTypeDrawer(SvgChart svgChart, ExcelLineChart chartType) : base(svgChart, chartType)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
            RenderItems.Add(groupItem);
            var isStacked = chartType.IsTypeStacked();
            var isPercentStacked = chartType.IsTypePercentStacked();
            var xValues = new List<List<object>>();
            var yValues = new List<List<object>>();
            int serCounter = 0;

            foreach (ExcelLineChartSerie serie in chartType.Series)
            {
                var yValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                var xValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                xValues.Add(xValue);
                yValues.Add(yValue);

                if (serie.HasDataLabel)
                {
                    var datalabel = new SvgChartSerieDataLabel(svgChart, serie.DataLabel, svgChart.Bounds, serie, xValue, yValue, serCounter);
                    serieDataLabels.Add(datalabel);
                }
                serCounter++;
            }

            if (chartType.IsTypeStacked())
            {
                SumSeries(yValues);
            }
            else if (chartType.IsTypePercentStacked())
            {
                ExcelChartAxisStandard.CalculateStacked100(yValues);
            }

            for(var i= 0; i < xValues.Count; i++)
            {
                var xSerie = xValues[i];
                var ySerie = yValues[i];
                var serie = chartType.Series[i];

                var dataPoints = new List<BoundingBox>();

                AddLine(chartType, serie, xSerie, ySerie, dataPoints);

                dataPointsPerSerie.Add(dataPoints);

                if (serie.HasDataLabel)
                {
                    for (int j = 0; j < dataPoints.Count; j++)
                    {
                        serieDataLabels[i].SetParentPoint(dataPoints[j], j);
                    }
                }
                if (serie.TrendLines.Count > 0)
                {
                    foreach (var trendline in serie.TrendLines)
                    {
                        var tr = new SvgTrendline(svgChart, trendline, xSerie, ySerie, chartType.UseSecondaryAxis);
                        Trendlines.Add(tr);
                    }
                }

            }
            RenderItems.Add(new SvgEndGroupItem(ChartRenderer, null));

            foreach (var dataLabel in serieDataLabels)
            {
                dataLabel.AppendRenderItems(RenderItems);
            }
        }

        private void SumSeries(List<List<object>> series)
        {
            for(var i=1;i < series.Count;i++)
            {
                for(var j=0;j < series[i].Count;j++)
                {
                    series[i][j] = ConvertUtil.GetValueDouble(series[i][j]) + ConvertUtil.GetValueDouble(series[i-1][j]);
                }
            }
        }
        private void AddLine(ExcelChart chartType, ExcelLineChartSerie serie, List<object> xValues, List<object> yValues, List<BoundingBox> dataPoints)
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
            var linePath = new SvgRenderPathItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
            var coords = new List<double>();
            var markerItems = new List<RenderItem>();

            for (var i = 0; i < yValues.Count; i++)
            {
                double x;
                if (xValues == null || xAxis.Axis.AxisType==eAxisType.Cat)
                {
                    x = (double)i;
                }
                else
                {
                    x = ConvertUtil.GetValueDouble(xValues[i], false, true);
                }

                var y = ConvertUtil.GetValueDouble(yValues[i], false, true);
                var xPos = xAxis.GetPositionInPlotarea(x);
                var yPos = yAxis.GetPositionInPlotarea(y);

                BoundingBox pt = null;

                if (double.IsNaN(yPos) == false)
                {
                    coords.Add(xPos / _svgChart.Plotarea.Rectangle.Bounds.Width);
                    coords.Add(yPos / _svgChart.Plotarea.Rectangle.Bounds.Height);

                    //Log point within chart coordinate system
                    pt = new BoundingBox(xPos, yPos, 0, 0);
                    pt.Parent = _svgChart.Plotarea.Bounds;
                }

                if (serie.HasMarker() && serie.Marker.Style != eMarkerStyle.None)
                {
                    float mx = (float)xPos;
                    float my = (float)yPos;
                    var ls = LineMarkerHelper.GetMarkerItem(_svgChart, serie, mx, my, false);
                    if ((serie.Marker.Style == eMarkerStyle.Plus || serie.Marker.Style == eMarkerStyle.X || serie.Marker.Style == eMarkerStyle.Star) &&
                        serie.Marker.Fill.IsEmpty == false)
                    {
                        markerItems.Add(LineMarkerHelper.GetMarkerBackground(_svgChart, serie, mx, my, false));
                    }
                    markerItems.Add(ls);

                    if (pt != null)
                    {
                        pt.Width = ls.Bounds.Width;
                        pt.Height = ls.Bounds.Height;
                        dataPoints.Add(pt);
                    }
                }
                else
                {

                    if (pt != null)
                    {
                        //Default values in excel
                        pt.Width = 5;
                        pt.Height = 5;
                        dataPoints.Add(pt);
                    }
                }
            }

            linePath.Commands.Add(new EPPlusImageRenderer.PathCommands(PathCommandType.Move, linePath, coords.ToArray()));
            linePath.SetDrawingPropertiesBorder(serie.Border, chartType.StyleManager.Style.SeriesLine.BorderReference.Color, true);
            linePath.SetDrawingPropertiesEffects(serie.Effect);
            linePath.FillColor = "none";    //No fill for line
            linePath.StrokeMiterLimit = 4;  //A much higher value of the miter limit, might cause the "spike" to get beyond the data point on the vertical scale..
            linePath.LineJoin = SvgLineJoin.Round; 
            RenderItems.Add(linePath);
            RenderItems.AddRange(markerItems);
        }
    }

}
