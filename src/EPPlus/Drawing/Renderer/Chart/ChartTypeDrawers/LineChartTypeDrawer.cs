using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class LineChartTypeDrawer : ChartTypeDrawer
    {
        List<List<object>> _xValues, _yValues;

        List<ChartSerieDataLabelRenderer> serieDataLabels = new List<ChartSerieDataLabelRenderer>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();
        internal override bool SupportsTrendlines => true;
        internal override bool SupportsErrorBars => true;
        internal LineChartTypeDrawer(ChartRenderer svgChart, ExcelLineChart chartType) : base(svgChart, chartType)
        {
            var isStacked = chartType.IsTypeStacked();
            var isPercentStacked = chartType.IsTypePercentStacked();
            int serCounter = 0;
            _xValues = new List<List<object>>();
            _yValues = new List<List<object>>();
            foreach (ExcelLineChartSerie serie in chartType.Series)
            {
                var yValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                var xValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                _xValues.Add(xValue);
                _yValues.Add(yValue);

                serCounter++;
            }

            if (chartType.IsTypeStacked())
            {
                SumSeries(_yValues);
            }
            else if (chartType.IsTypePercentStacked())
            {
                ExcelChartAxisStandard.CalculateStacked100(_yValues);
            }

            CreateTrendlines(chartType, _xValues, _yValues);
            CreateErrorBars(chartType, _xValues, _yValues);
        }
        internal override void DrawSeries()
        {
            var lct = (ExcelLineChart)_chartType;
            for (var i = 0; i < _xValues.Count; i++)
            {
                var xSerie = _xValues[i];
                var ySerie = _yValues[i];
                var serie = lct.Series[i];

                if (serie.HasDataLabel)
                {
                    var datalabel = new ChartSerieDataLabelRenderer(ChartRenderer, serie.DataLabel, ChartRenderer.Bounds, serie, xSerie, ySerie, i);
                    serieDataLabels.Add(datalabel);
                }


                var dataPoints = new List<BoundingBox>();
                AddLine(_chartType, serie, xSerie, ySerie, dataPoints);

                dataPointsPerSerie.Add(dataPoints);

                if (serie.HasDataLabel)
                {
                    for (int j = 0; j < dataPoints.Count; j++)
                    {
                        serieDataLabels[i].SetParentPoint(dataPoints[j], j);
                    }
                }
            }


            //Append Trendline render items.
            foreach (var tr in Trendlines)
            {
                tr.CreateRenderCoordinatesAndDatalabel();
                tr.AppendRenderItems(SeriesRenderItems);
            }

            //Datalabels use the chart area as parent as they can be positioned on the entire chart.

            //Add data labels for trendlines after the trendline has been rendered, to ensure they are on top of the line.
            foreach (var tr in Trendlines)
            {
                if (tr.DataLabel != null)
                {
                    tr.DataLabel.AppendRenderItems(ChartAreaRenderItems);
                }
            }

            //Date series labels
            foreach (var dataLabel in serieDataLabels)
            {
                dataLabel.AppendRenderItems(ChartAreaRenderItems);
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
            ChartAxisRenderer yAxis, xAxis;
            if (chartType.UseSecondaryAxis)
            {
                yAxis = ChartRenderer.SecondVerticalAxis;
                xAxis = ChartRenderer.SecondHorizontalAxis;
                if(xAxis.Axis.Deleted && xAxis.Values==null)
                {
                    xAxis = ChartRenderer.HorizontalAxis;
                }
            }
            else
            {
                yAxis = ChartRenderer.VerticalAxis;
                xAxis = ChartRenderer.HorizontalAxis;
            }
            var linePath = new PathRenderItem(ChartRenderer.Plotarea.Rectangle.Bounds);
            var coords = new List<double>();
            var markerItems = new List<RenderItem>();
            var errorBars = new List<RenderItem>();

            var hasMarker = serie.HasMarker() && serie.Marker.Style != eMarkerStyle.None;
            var hasErrorBars = serie.HasErrorBars();
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
                    coords.Add(xPos);
                    coords.Add(yPos);

                    //Log point within chart coordinate system
                    pt = new BoundingBox(xPos, yPos, 0, 0);
                    pt.Parent = ChartRenderer.Plotarea.Rectangle.Bounds;
                }
                if(hasErrorBars)
                {
                    errorBars.AddRange(ErrorBars.GetErrorBarRenderItem(i, xAxis, yAxis, x, y, xPos, yPos));
                }
                if (hasMarker)
                {
                    float mx = (float)xPos;
                    float my = (float)yPos;
                    var ls = LineMarkerHelper.GetMarkerItem(ChartRenderer, serie, mx, my, false);
                    if ((serie.Marker.Style == eMarkerStyle.Plus || serie.Marker.Style == eMarkerStyle.X || serie.Marker.Style == eMarkerStyle.Star) &&
                        serie.Marker.Fill.IsEmpty == false)
                    {
                        markerItems.Add(LineMarkerHelper.GetMarkerBackground(ChartRenderer, serie, mx, my, false));
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

            linePath.Commands.Add(new PathCommands(PathCommandType.Move, coords.ToArray()));
            linePath.SetDrawingPropertiesBorder(ChartRenderer.Theme, serie.Border, chartType.StyleManager.Style?.SeriesLine.BorderReference.Color, true);
            linePath.SetDrawingPropertiesEffects(ChartRenderer.Theme, serie.Effect);
            linePath.FillColor = "none";    //No fill for line
            linePath.StrokeMiterLimit = 4;  //A much higher value of the miter limit, might cause the "spike" to get beyond the data point on the vertical scale..
            linePath.LineJoin = LineJoin.Round;
            SeriesRenderItems.Add(linePath);
            SeriesRenderItems.AddRange(markerItems);
            SeriesRenderItems.AddRange(errorBars);
        }


        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.AddRange(ChartAreaRenderItems);
            SeriesRenderItems.ForEach(x=> ChartRenderer.Plotarea.Group.AddChildItem(x));
        }
    }

}
