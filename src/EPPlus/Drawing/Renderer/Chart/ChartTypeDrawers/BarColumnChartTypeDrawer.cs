using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class BarColumnChartTypeDrawer : ChartTypeDrawer
    {
        List<List<object>> _catValues, _valValues;
        List<ChartSerieDataLabelRenderer> serieDataLabels = new List<ChartSerieDataLabelRenderer>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();
        internal override bool SupportsTrendlines => true;
        internal override bool SupportsErrorBars => true;
        internal BarColumnChartTypeDrawer(ChartRenderer svgChart, ExcelBarChart chartType) : base(svgChart, chartType)
        {
            _catValues = new List<List<object>>();
            _valValues = new List<List<object>>();

            int serCounter = 0;

            foreach (ExcelBarChartSerie serie in chartType.Series)
            {
                List<object> valValue,catValue;
                valValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                catValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                _catValues.Add(catValue);
                _valValues.Add(valValue);
            }

            if(chartType.IsTypeStacked())
            {
                SumSeries(_valValues);
            }
            else if (chartType.IsTypePercentStacked())
            {
                ExcelChartAxisStandard.CalculateStacked100(_valValues);
            }

            CreateTrendlines(chartType, _catValues, _valValues);
            CreateErrorBars(chartType, _catValues, _valValues);
        }

        internal override void DrawSeries()
        {
            var isBar = _chartType.IsTypeBar();
            var count = Math.Min(_catValues.Count, _valValues.Count);
            for (var i = 0; i < _catValues.Count; i++)
            {
                var serie = (ExcelBarChartSerie)_chartType.Series[i];

                var dataPoints = new List<BoundingBox>();

                //Add the bar or column.
                AddBar((ExcelBarChart)_chartType, serie, _catValues, _valValues, dataPoints, count, i);

                dataPointsPerSerie.Add(dataPoints);

                int serCounter = 0;

                var isColumn = ((ExcelBarChart)_chartType).IsTypeColumn();

                if (serie.HasDataLabel)
                {
                    var datalabel = new ChartSerieDataLabelRenderer(ChartRenderer, serie.DataLabel, ChartRenderer.Bounds, serie, _catValues[i], _valValues[i], serCounter++);
                    serieDataLabels.Add(datalabel);

                    for (int j = 0; j < dataPoints.Count; j++)
                    {
                        //Initialize transforms
                        Transform basePoint = new Transform();
                        Transform endPoint = new Transform();
                        basePoint.Parent = dataPoints[j];
                        endPoint.Parent = dataPoints[j];

                        if (isColumn == true)
                        {
                            var middleRight = dataPoints[j].Left + (dataPoints[j].Width / 2);
                            basePoint.Position = new Vector2(middleRight, chartBaseY);

                            if (chartBaseY > dataPoints[j].Top)
                            {
                                endPoint.Position = new Vector2(middleRight, chartBaseY - dataPoints[j].Height);
                            }
                            else
                            {
                                endPoint.Position = new Vector2(middleRight, chartBaseY + dataPoints[j].Height);
                            }

                            serieDataLabels[i].SetDimensions(j, basePoint, endPoint);
                        }
                        else
                        {
                            var middleHeight = dataPoints[j].Top + (dataPoints[j].Height / 2);


                            basePoint.Position = new Vector2(basePoint.Position.X, middleHeight);
                            endPoint.Position = new Vector2(basePoint.Position.X + dataPoints[j].Width, middleHeight);

                            serieDataLabels[i].SetDimensions(j, basePoint, endPoint);
                        }
                    }

                    serCounter++;
                }
            }

            foreach (var tr in Trendlines)
            {
                tr.CreateRenderCoordinatesAndDatalabel();
                tr.AppendRenderItems(SeriesRenderItems);
            }

            //Add data labels for trendlines after the trendline has been rendered, to ensure they are on top of the line.
            foreach (var tr in Trendlines)
            {
                if (tr.DataLabel != null)
                {
                    tr.DataLabel.AppendRenderItems(ChartAreaRenderItems);
                }
            }

            foreach (var dataLabel in serieDataLabels)
            {
                dataLabel.AppendRenderItems(ChartAreaRenderItems);
            }
        }

        double chartBaseY = double.NaN;

        private void AddBar(ExcelBarChart chartType, ExcelBarChartSerie serie, List<List<object>> catSeries, List<List<object>> valSeries, List<BoundingBox> dataPoints, int seriesCount, int position)
        {
            GetAxis(chartType, out var yAxis, out var xAxis);
            ChartAxisRenderer valAx, catAx;

            var isColumn = chartType.IsTypeColumn();

            if (isColumn)
            {
                catAx = xAxis;
                valAx = yAxis;
            }
            else
            {
                catAx = yAxis;
                valAx = xAxis;
            }

            var catValues = catSeries[position];
            var valValues = valSeries[position];
            
            var yWidth = (isColumn ? ChartRenderer.Plotarea.Rectangle.Width : ChartRenderer.Plotarea.Rectangle.Height);

            var slotSize = valValues.Count;
            var gapPercent = chartType.GapWidth / 100D;     // Gap width between bars/columns in percent
            var overlapPercent = chartType.Overlap / 100D;  // Overlap  between bars/columns in percent            
            var slotWidth = yWidth / slotSize;
            var clusterWidth = slotWidth * 100 / (100 + chartType.GapWidth);
            var step = 1 - overlapPercent;
            var barWidth = slotWidth / (1 + (seriesCount - 1) * step + gapPercent);
            var halfGap = (barWidth * gapPercent) / 2;

            if (catAx.Axis.Crosses == eCrosses.AutoZero)
            {
                chartBaseY = valAx.GetPositionInPlotarea(valAx.Min <= 0 ? 0D : valAx.Min, true);
            }
            else if (catAx.Axis.Crosses == eCrosses.Min)
            {
                chartBaseY = valAx.GetPositionInPlotarea(valAx.Min, true);
            }
            else
            {
                chartBaseY = valAx.GetPositionInPlotarea(valAx.Max, true);
            }

            var isStacked = chartType.IsTypeStacked();
            var isStacked100 = chartType.IsTypePercentStacked();
            for (var i = 0; i < valValues.Count; i++)
            {
                double x;
                if (catValues == null || catAx.Axis.AxisType == eAxisType.Cat)
                {
                    if (isColumn)
                    {
                        x = (double)i;
                    }
                    else
                    {
                        x = valValues.Count - i - 1;
                    }
                }
                else
                {
                    x = ConvertUtil.GetValueDouble(catValues[i], false, true);
                }

                var y = ConvertUtil.GetValueDouble(valValues[i], false, true);
                
                var rect = new RectRenderItem(ChartRenderer.Plotarea.Rectangle.Bounds);
                var yPos = valAx.GetPositionInPlotarea(y);

                if (isColumn)
                {
                    var xPos = catAx.GetPositionInPlotarea(x, true) + halfGap + position * barWidth * step;
                    rect.Left = xPos;
                    rect.Width = barWidth;
                }
                else
                {
                    var xPos = catAx.GetPositionInPlotarea(x, true) + halfGap + (seriesCount - position - 1) * barWidth * step;
                    rect.Top = xPos;
                    rect.Height = barWidth;
                }

                if (position > 0 && (isStacked || isStacked100))
                {
                    //Stacked
                    var pYValues = valSeries[position - 1];
                    var pAxisEnd = ConvertUtil.GetValueDouble(pYValues[i]);
                    var yPrevPos = yAxis.GetPositionInPlotarea(pAxisEnd, false);
                    if (isColumn)
                    {
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
                            rect.Left = yPrevPos;
                            rect.Width = yPos - yPrevPos;
                        }
                        else
                        {
                            rect.Left = yPos;
                            rect.Width = yPrevPos - yPos;
                        }
                    }
                }
                else
                {
                    if (isColumn)
                    {
                        if (y < 0)
                        {
                            rect.Top = chartBaseY;
                            rect.Height = yPos - chartBaseY;
                        }
                        else
                        {
                            rect.Top = yPos;
                            rect.Height = chartBaseY - yPos;
                        }
                    }
                    else
                    {
                        if (y < 0)
                        {
                            rect.Left = yPos;
                            rect.Width = chartBaseY - yPos;
                        }
                        else
                        {
                            rect.Left = chartBaseY;
                            rect.Width = yPos - chartBaseY;
                        }
                    }
                }
                rect.SetDrawingPropertiesFill(ChartRenderer.Theme, serie.Fill, chartType.StyleManager.Style?.SeriesAxis.FillReference.Color);
                rect.SetDrawingPropertiesBorder(ChartRenderer.Theme, serie.Border, chartType.StyleManager.Style?.SeriesAxis.BorderReference.Color, true);
                rect.SetDrawingPropertiesEffects(ChartRenderer.Theme, serie.Effect);

                dataPoints.Add(rect.Bounds);

                SeriesRenderItems.Add(rect);
            }
        }
        private void GetAxis(ExcelBarChart chartType, out ChartAxisRenderer yAxis, out ChartAxisRenderer xAxis)
        {
            if (chartType.UseSecondaryAxis)
            {
                yAxis = ChartRenderer.SecondVerticalAxis;
                xAxis = ChartRenderer.SecondHorizontalAxis;
                if (xAxis.Axis.Deleted && xAxis.Values == null)
                {
                    xAxis = ChartRenderer.HorizontalAxis;
                }
            }
            else
            {
                yAxis = ChartRenderer.VerticalAxis;
                xAxis = ChartRenderer.HorizontalAxis;
            }
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.AddRange(ChartAreaRenderItems);
            SeriesRenderItems.ForEach(x => ChartRenderer.Plotarea.Group.AddChildItem(x));
        }
    }
}
