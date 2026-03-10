using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using static OfficeOpenXml.ExcelErrorValue;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class LineChartTypeDrawer : ChartTypeDrawer
    {
        internal LineChartTypeDrawer(SvgChart svgChart, ExcelChart chartType) : base(svgChart, chartType)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
            RenderItems.Add(groupItem);
            var isStacked = Chart.IsTypeStacked();
            var isPercentStacked = Chart.IsTypePercentStacked();
            var xValues = new List<List<object>>();
            var yValues = new List<List<object>>();
            var serieDataLabels = new List<SvgChartSerieDataLabel>();

            foreach (ExcelLineChartSerie serie in chartType.Series)
            {
                var yValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                var xValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                xValues.Add(xValue);
                yValues.Add(yValue);

                if (serie.HasDataLabel)
                {
                    var datalabel = new SvgChartSerieDataLabel(svgChart, serie.DataLabel, svgChart.Plotarea.Bounds);
                    serieDataLabels.Add(datalabel);
                }
            }

            if (Chart.IsTypeStacked())
            {
                SumSeries(yValues);
            }
            else if (Chart.IsTypePercentStacked())
            {
                ExcelChartAxisStandard.CalculateStacked100(yValues);
            }

            for(var i= 0; i < xValues.Count; i++)
            {
                var xSerie = xValues[i];
                var ySerie = yValues[i];
                var serie = (ExcelLineChartSerie)chartType.Series[i];
                AddLine(chartType, serie, xSerie, ySerie);
            }

            foreach(var dataLabel in serieDataLabels)
            {
                dataLabel.AppendRenderItems(RenderItems);
            }

            RenderItems.Add(new SvgEndGroupItem(ChartRenderer, null));
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
        private void AddLine(ExcelChart chartType, ExcelLineChartSerie serie, List<object> xValues, List<object> yValues)
        {
            var xAxis = _svgChart.HorizontalAxis;
            SvgChartAxis yAxis;
            if (chartType.UseSecondaryAxis)
            {
                yAxis = _svgChart.SecondVerticalAxis;
            }
            else
            {
                yAxis = _svgChart.VerticalAxis;
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

                if (double.IsNaN(yPos) == false)
                {
                    coords.Add(xPos / _svgChart.Plotarea.Rectangle.Bounds.Width);
                    coords.Add(yPos / _svgChart.Plotarea.Rectangle.Bounds.Height);
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
                }
            }

            linePath.Commands.Add(new EPPlusImageRenderer.PathCommands(PathCommandType.Move, linePath, coords.ToArray()));
            linePath.SetDrawingPropertiesBorder(serie.Border, chartType.StyleManager.Style.SeriesLine.BorderReference.Color, true);
            linePath.SetDrawingPropertiesEffects(serie.Effect);
            linePath.FillColor = "none"; // No fill for line
            RenderItems.Add(linePath);
            RenderItems.AddRange(markerItems);
        }
    }

}
