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
            var catValues = new List<List<object>>();
            var valValues = new List<List<object>>();
            int serCounter = 0;

            foreach (ExcelBarChartSerie serie in chartType.Series)
            {
                List<object> valValue,catValue;
                //if (chartType.IsTypeColumn())
                //{
                    valValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                    catValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);
                //}
                //else
                //{
                //    catValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                //    valValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);
                //}

                catValues.Add(catValue);
                valValues.Add(valValue);

                //if (serie.HasDataLabel)
                //{
                //    var datalabel = new SvgChartSerieDataLabel(svgChart, serie.DataLabel, svgChart.Bounds, serie, catValue, valValue, serCounter);
                //    serieDataLabels.Add(datalabel);
                //}
                serCounter++;
            }

            if(chartType.IsTypeStacked())
            {
                SumSeries(valValues);
            }
            else if (chartType.IsTypePercentStacked())
            {
                ExcelChartAxisStandard.CalculateStacked100(valValues);
            }

            var count = Math.Min(catValues.Count, valValues.Count);
            for (var i= 0; i < catValues.Count; i++)  
            {
                var serie = (ExcelBarChartSerie)chartType.Series[i];

                var dataPoints = new List<BoundingBox>();

                //Add the bar or column.
                AddBar(chartType, serie, catValues, valValues, dataPoints, count, i);

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
        private void AddBar(ExcelBarChart chartType, ExcelBarChartSerie serie, List<List<object>> catSeries, List<List<object>> valSeries, List<BoundingBox> dataPoints, int seriesCount, int position)
        {
            GetAxis(chartType, out var yAxis, out var xAxis);
            SvgChartAxis valAx, catAx;

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
            
            var yWidth = (isColumn ? _svgChart.Plotarea.Rectangle.Width : _svgChart.Plotarea.Rectangle.Height);

            var slotSize = valValues.Count;
            var gapPercent = chartType.GapWidth / 100D;     //Gap width between bars/columns in percent
            var overlapPercent = chartType.Overlap / 100D;  //Overlap  between bars/columns in percent            
            var slotWidth = yWidth / slotSize;
            var clusterWidth = slotWidth * 100 / (100 + chartType.GapWidth);
            var step = 1 - overlapPercent;
            var barWidth = slotWidth / (1 + (seriesCount - 1) * step + gapPercent);
            var halfGap = (barWidth * gapPercent) / 2;

            double yAxisStart;
            if (yAxis.Axis.Crosses == eCrosses.AutoZero)
            {
                yAxisStart = valAx.GetPositionInPlotarea(valAx.Min <= 0 ? 0D : valAx.Min, true);
            }
            else if (yAxis.Axis.Crosses == eCrosses.Min)
            {
                yAxisStart = valAx.GetPositionInPlotarea(valAx.Min, true);
            }
            else
            {
                yAxisStart = valAx.GetPositionInPlotarea(valAx.Max, true);
            }

            var isStacked = chartType.IsTypeStacked();
            var isStacked100 = chartType.IsTypePercentStacked();
            for (var i = 0; i < valValues.Count; i++)
            {
                double x;
                if (catValues == null || catAx.Axis.AxisType == eAxisType.Cat)
                {
                    x = (double)i;
                }
                else
                {
                    x = ConvertUtil.GetValueDouble(catValues[i], false, true);
                }

                var y = ConvertUtil.GetValueDouble(valValues[i], false, true);
                
                var rect = new SvgRenderRectItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
                var yPos = valAx.GetPositionInPlotarea(y);

                if (isColumn)
                {
                    var xPos = catAx.GetPositionInPlotarea(x, true) + halfGap + position * barWidth * step;
                    rect.Left = xPos;
                    rect.Width = barWidth;
                }
                else
                {
                    //var xPos = catAx.GetPositionInPlotarea(x, true) - halfGap - (position + 1) * barWidth * step;
                    var xPos = catAx.GetPositionInPlotarea(x, true) + halfGap + position * barWidth * step;
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
                            rect.Top = yAxisStart;
                            rect.Height = yPos - yAxisStart;
                        }
                        else
                        {
                            rect.Top = yPos;
                            rect.Height = yAxisStart - yPos;
                        }
                    }
                    else
                    {
                        if (y < 0)
                        {
                            rect.Left = yPos;
                            rect.Width = yAxisStart - yPos;
                        }
                        else
                        {
                            rect.Left = yAxisStart;
                            rect.Width = yPos - yAxisStart;
                        }
                    }
                }
                rect.SetDrawingPropertiesFill(serie.Fill, chartType.StyleManager.Style.SeriesAxis.FillReference.Color);
                rect.SetDrawingPropertiesBorder(serie.Border, chartType.StyleManager.Style.SeriesAxis.BorderReference.Color, true);
                rect.SetDrawingPropertiesEffects(serie.Effect);
                RenderItems.Add(rect);
            }
        }
        private void AddColumn(ExcelBarChart chartType, ExcelBarChartSerie serie, List<List<object>> xSeries, List<List<object>> ySeries, List<BoundingBox> dataPoints, int seriesCount, int position)
        {
            SvgChartAxis yAxis, xAxis;
            GetAxis(chartType, out yAxis, out xAxis);

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
                if (position > 0 && (isStacked || isStacked100))
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

        private void GetAxis(ExcelBarChart chartType, out SvgChartAxis yAxis, out SvgChartAxis xAxis)
        {
            if (chartType.UseSecondaryAxis)
            {
                yAxis = _svgChart.SecondVerticalAxis;
                xAxis = _svgChart.SecondHorizontalAxis;
                if (xAxis.Axis.Deleted && xAxis.Values == null)
                {
                    xAxis = _svgChart.HorizontalAxis;
                }
            }
            else
            {
                yAxis = _svgChart.VerticalAxis;
                xAxis = _svgChart.HorizontalAxis;
            }
        }
    }

}
