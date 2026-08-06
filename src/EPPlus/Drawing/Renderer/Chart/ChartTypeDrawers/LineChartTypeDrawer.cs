using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using static OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingConstants;

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
                if(xValue==null)
                {
                    //No x-axis. Create serie from 1..y-items.
                    xValue = yValue.Select((x, index) => (object)(double)(index + 1)).ToList();
                }
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
        List<LineRenderItem> _dropLines=null;
        private void CreateDropLine(ExcelLineChart chartType, List<double> coords)
        {
            if (chartType.DropLine == null) return;
            _dropLines=new List<LineRenderItem>();
            for (var i = 0; i < coords.Count; i += 2)
            {
                var x = coords[i];
                var yTop = coords[i+1];
                var catAxis = chartType.UseSecondaryAxis ? ChartRenderer.SecondHorizontalAxis : ChartRenderer.HorizontalAxis;
                var valAxis = chartType.UseSecondaryAxis ? ChartRenderer.SecondVerticalAxis : ChartRenderer.VerticalAxis;
                var yBottom = GetAxisBaseY(catAxis, valAxis);
                var bb = ChartRenderer.Plotarea.Group.Bounds;
                var dl = new LineRenderItem(bb)
                {
                    X1 = x,
                    X2 = x,
                    Y1 = yTop,
                    Y2 = yBottom,                    
                };
                dl.Bounds.Name = $"DropLine {i/2 + 1}";
                dl.SetDrawingPropertiesBorder(ChartRenderer.Theme, chartType.DropLine.Border, chartType.StyleManager.Style?.DropLine.BorderReference.Color, true, DefaultBorderColor, 1.5,DrawingRenderer.UserSpaceSettings.UserSpaceOnUse_Parent);
                dl.SetDrawingPropertiesEffects(ChartRenderer.Theme, chartType.DropLine.Effect);
                
                _dropLines.Add(dl);
            }
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
                AddLine(_chartType.As.Chart.LineChart, serie, xSerie, ySerie, dataPoints);

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
        private new void SumSeries(List<List<object>> series)
        {
            for(var i=1;i < series.Count;i++)
            {
                for(var j=0;j < series[i].Count;j++)
                {
                    series[i][j] = ConvertUtil.GetValueDouble(series[i][j]) + ConvertUtil.GetValueDouble(series[i-1][j]);
                }
            }
        }
        private void AddLine(ExcelLineChart chartType, ExcelLineChartSerie serie, List<object> xValues, List<object> yValues, List<BoundingBox> dataPoints)
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
            var dataPointOverrides = new List<LineRenderItem>();
            var coords = new List<double>();
            var markerItems = new List<RenderItem>();
            var errorBars = new List<RenderItem>();

            var hasMarker = serie.HasMarker() && serie.Marker.Style != eMarkerStyle.None;
            var hasErrorBars = serie.HasErrorBars();
            for (var i = 0; i < yValues.Count; i++)
            {
                double x;
                if (xValues == null || (xAxis.Axis.AxisType==eAxisType.Cat && xValues.Count>0 && !(xValues[0] is double)))
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
                    SetMarker(serie, serie.Marker, dataPoints, markerItems, xPos, yPos, pt);
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
                //Draw individual formatted lines between points, if the previous point has a different formatting than the current point.
                if (i > 0 && serie.DataPoints.ContainsKey(i))
                {
                    var dp = serie.DataPoints[i];
                    var lineDp = new LineRenderItem(ChartRenderer.Plotarea.Rectangle.Bounds);
                    lineDp.X1 = coords[coords.Count - 4];
                    lineDp.Y1 = coords[coords.Count - 3];
                    lineDp.X2 = xPos;
                    lineDp.Y2 = yPos;

                    if(dp.HasMarker())
                    {
                        if(hasMarker==false)
                        {
                            SetMarker(serie, dp.Marker, dataPoints, markerItems, xPos, yPos, pt);
                        }
                        else
                        {
                            var mi = markerItems[markerItems.Count - 1];
                            markerItems[i].SetDrawingPropertiesFill(ChartRenderer.Theme, dp.Marker.Fill, chartType.StyleManager.Style?.DataPointMarker.FillReference.Color);
                            markerItems[i].SetDrawingPropertiesBorder(ChartRenderer.Theme, dp.Marker.Border, chartType.StyleManager.Style?.DataPointMarker.FillReference.Color, serie.Border.Fill.Style != eFillStyle.NoFill);
                        }
                    }
                    lineDp.SetDrawingPropertiesBorder(ChartRenderer.Theme, dp.Border, chartType.StyleManager.Style?.SeriesLine.BorderReference.Color, true, DefaultBorderColor, 3);
                    lineDp.SetDrawingPropertiesEffects(ChartRenderer.Theme, dp.Effect);
                    dataPointOverrides.Add(lineDp);
                }
            }
            
            CreateDropLine(chartType, coords);

            linePath.Commands.Add(new PathCommands(PathCommandType.Move, coords.ToArray()));
            linePath.SetDrawingPropertiesBorder(ChartRenderer.Theme, serie.Border, chartType.StyleManager.Style?.SeriesLine.BorderReference.Color, true, DefaultBorderColor, 3);
            linePath.SetDrawingPropertiesEffects(ChartRenderer.Theme, serie.Effect);
            linePath.FillColor = "none";    //No fill for line
            linePath.StrokeMiterLimit = 4;  //A much higher value of the miter limit, might cause the "spike" to get beyond the data point on the vertical scale..
            linePath.LineJoin = LineJoin.Round;
            SeriesRenderItems.Add(linePath);
            SeriesRenderItems.AddRange(dataPointOverrides);
            SeriesRenderItems.AddRange(markerItems);
            if(_dropLines!=null) SeriesRenderItems.AddRange(_dropLines);
            SeriesRenderItems.AddRange(errorBars);
        }

        private void SetMarker(ExcelLineChartSerie serie, ExcelChartMarker marker, List<BoundingBox> dataPoints, List<RenderItem> markerItems, double xPos, double yPos, BoundingBox pt)
        {
            float mx = (float)xPos;
            float my = (float)yPos;
            var ls = LineMarkerHelper.GetMarkerItem(ChartRenderer, serie, marker, mx, my, false);
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

        internal override Color? DefaultBorderColor => ChartRenderer.Theme.ColorScheme.Accent1.GetColor();
        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.AddRange(ChartAreaRenderItems);
            SeriesRenderItems.ForEach(x=> ChartRenderer.Plotarea.Group.AddChildItem(x));
        }
    }

}
