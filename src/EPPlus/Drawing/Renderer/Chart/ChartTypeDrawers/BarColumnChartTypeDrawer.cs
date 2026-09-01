using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class BarColumnChartTypeDrawer : ChartTypeDrawer
    {
        List<List<object>> _catValues, _valValues, _origValValues;
        List<ChartSerieDataLabelRenderer> serieDataLabels = new List<ChartSerieDataLabelRenderer>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();
        internal override bool SupportsTrendlines => true;
        internal override bool SupportsErrorBars => true;
        internal override bool SupportsDataTable => true;

        internal BarColumnChartTypeDrawer(ChartRenderer svgChart, ExcelBarChart chartType) : base(svgChart, chartType)
        {
            _catValues = new List<List<object>>();
            _valValues = new List<List<object>>();
            _origValValues = new List<List<object>>();

            int serCounter = 0;

            foreach (ExcelBarChartSerie serie in chartType.Series)
            {
                List<object> valValue, catValue;
                valValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                catValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                _catValues.Add(catValue);
                _valValues.Add(valValue);

                List<object> origList = new List<object>();
                if(valValue != null)
                {
                    for (int i = 0; i < valValue.Count; i++)
                    {
                        origList.Add(valValue[i]);
                    }
                    //Will not be summed
                    _origValValues.Add(origList);
                }
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
            for (var i = 0; i < count; i++)
            {
                var serie = (ExcelBarChartSerie)_chartType.Series[i];

                var dataPoints = new List<BoundingBox>();

                //Add the bar or column.
                AddBar((ExcelBarChart)_chartType, serie, _catValues, _valValues, dataPoints, count, i);

                dataPointsPerSerie.Add(dataPoints);

                int serCounter = 0;

                var isColumn = ((ExcelBarChart)_chartType).IsTypeColumn();

                if (serie.HasDataLabel && serie.DataLabel != null)
                {
                    var datalabel = new ChartSerieDataLabelRenderer(ChartRenderer, serie.DataLabel, ChartRenderer.Bounds, serie, _catValues[i], _origValValues[i], serCounter++);
                    serieDataLabels.Add(datalabel);

                    for (int j = 0; j < dataPoints.Count; j++)
                    {
                        //var parentHolder = dataPoints[j].Parent;

                        var globalDPBounds = dataPoints[j].GetGlobalBoundingbox();

                        //Initialize transforms
                        Transform basePoint = new Transform();
                        Transform endPoint = new Transform();
                        basePoint.Parent = dataPoints[j];
                        endPoint.Parent = dataPoints[j];

                        if (isColumn == true)
                        {
                            var middleRight = globalDPBounds.Left + (globalDPBounds.Width / 2);

                            if (chartBaseY <= globalDPBounds.Top)
                            {
                                //We are a negative column 
                                // ----- Base-Axis
                                //  |_| Col
                                basePoint.Position = new Vector2(middleRight, globalDPBounds.Top);
                                endPoint.Position = new Vector2(middleRight, globalDPBounds.Bottom);
                            }
                            else
                            {
                                //We are a positive column 
                                //   _
                                //  | |  Col
                                // ----- Base-Axis
                                basePoint.Position = new Vector2(middleRight, globalDPBounds.Bottom);
                                endPoint.Position = new Vector2(middleRight, globalDPBounds.Top);
                            }

                            datalabel.SetDimensions(j, basePoint, endPoint);
                        }
                        else
                        {
                            var middleHeight = globalDPBounds.Top + (globalDPBounds.Height / 2);
                            basePoint.Position = new Vector2(chartBaseY, middleHeight);
                            if (chartBaseY > globalDPBounds.Left)
                            {
                                endPoint.Position = new Vector2(chartBaseY - globalDPBounds.Width, middleHeight);
                            }
                            else
                            {
                                endPoint.Position = new Vector2(globalDPBounds.Left + globalDPBounds.Width, middleHeight);
                            }

                            datalabel.SetDimensions(j, basePoint, endPoint);
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

            var hasErrorBars = serie.HasErrorBars();

            var catValues = catSeries[position];
            var valValues = valSeries[position];
            
            var yWidth = (isColumn ? ChartRenderer.Plotarea.Rectangle.Width : ChartRenderer.Plotarea.Rectangle.Height);
            double slotSize;
            //if(catAx.IsDateScale)
            ////{
            //    slotSize = (catAx.Max - catAx.Min) / catAx.MajorUnit + 1; //valValues.Count
            //}
            //else
            //{
                 slotSize = catAx.Max - catAx.Min+1;
            //}
            //var slotSize = (catAx.Max - catAx.Min) / catAx.MajorUnit + 1; //valValues.Count;
            //Gapwidth has a default value of 150% See ECMA-376 Part 1 page 4063:
            //"<xsd:complexType name="CT_GapAmount">286 <xsd:attribute name="val" type="ST_GapAmount" default="150%"/>287 </xsd:complexType>"
            var gapWidth = chartType.GapWidth == int.MinValue ? 150 : chartType.GapWidth;
            var gapPercent = gapWidth / 100D;               // Gap width between bars/columns in percent
            var overlapPercent = chartType.Overlap / 100D;  // Overlap  between bars/columns in percent            
            var slotWidth = yWidth / slotSize;
            var clusterWidth = slotWidth * 100 / (100 + gapWidth);
            var step = 1 - overlapPercent;
            var barWidth = slotWidth / (1 + (seriesCount - 1) * step + gapPercent);
            var halfGap = (barWidth * gapPercent) / 2;

            chartBaseY = GetAxisBaseY(catAx, valAx);

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
                    if (isColumn)
                    {
                        x = ConvertUtil.GetValueDouble(catValues[i], false, true);
                    }
                    else
                    {
                        var ix = valValues.Count - i - 1;
                        if (ix < catValues.Count)
                        {
                            x = ConvertUtil.GetValueDouble(catValues[ix], false, true);
                        }
                        else
                        {
                            x = 0;
                        }
                    }
                }

                var y = ConvertUtil.GetValueDouble(valValues[i], false, true);
                
                var rect = new RectRenderItem(ChartRenderer.Plotarea.Rectangle.Bounds);
                var yPos = valAx.GetPositionInPlotarea(y);
                double xPos;
                if (isColumn)
                {
                    xPos = catAx.GetPositionInPlotarea(x, true) + halfGap + position * barWidth * step;
                    rect.Left = xPos;
                    rect.Width = barWidth;
                }
                else
                {
                    xPos = catAx.GetPositionInPlotarea(x, true) + halfGap + (seriesCount - position - 1) * barWidth * step;
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

                //rect.SetDrawingPropertiesFill(ChartRenderer.Theme, serie.Fill, chartType.StyleManager.Style?.SeriesAxis.FillReference.Color, false, DefaultFillColor);
                //rect.SetDrawingPropertiesBorder(ChartRenderer.Theme, serie.Border, chartType.StyleManager.Style?.SeriesAxis.BorderReference.Color, serie.Border.IsEmpty==false && serie.Border.Width>0, DefaultBorderColor, 1.5, false);
                if (i >= 0 && serie.DataPoints.ContainsKey(i))
                {
                    var dp = serie.DataPoints[i];
                    SetFillDataPoint(Chart, serie, i, rect, dp, Chart.StyleManager.Style?.SeriesLine);
                }
                else
                {
                    SetFillSerie(Chart, chartType, serie, position, i, rect);
                }

                rect.SetDrawingPropertiesEffects(ChartRenderer.Theme, serie.Effect);

                dataPoints.Add(rect.Bounds);

                SeriesRenderItems.Add(rect);

                if (hasErrorBars)
                {
                    if(isColumn)
                    {
                        //x += barWidth / 2;
                        xPos += rect.Width / 2;
                    }
                    else
                    {
                        //y += barWidth / 2;
                        xPos += rect.Height / 2;
                    }
                    SeriesRenderItems.AddRange(ErrorBars.GetErrorBarRenderItem(i, catAx, valAx, x, y, xPos, yPos));
                }
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
        internal override Color? DefaultFillColor => ChartRenderer.Theme.ColorScheme.Accent1.GetColor();
        internal override Color? DefaultBorderColor => null;
    }
}
