using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class LineChartTypeDrawer : ChartTypeDrawer
    {
        internal LineChartTypeDrawer(SvgChart svgChart, ExcelChart chartType) : base(svgChart, chartType)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
            RenderItems.Add(groupItem);
            var isStacked = Chart.IsTypeStacked();
            List<object> pXValues = null, pYValues = null;
            foreach (ExcelLineChartSerie serie in chartType.Series)
            {
                var yValues = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                var xValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);
                if(isStacked)
                {
                    AddLine(chartType, serie, xValues, SumLists(yValues, pYValues));
                }
                else
                {
                    AddLine(chartType, serie, xValues, yValues);
                }

                if (isStacked)
                {
                    pYValues = yValues;
                }
            }
            RenderItems.Add(new SvgEndGroupItem(ChartRenderer, null));
        }
        private List<object> SumLists(List<object> l1, List<object> l2)
        {
            if (l2 == null) return l1;
            var sumList = new List<object>(l1.Count);
            for (int i = 0; i < Math.Min(l1.Count, l2.Count); i++)
            {
                sumList.Add(ConvertUtil.GetValueDouble(l1[i]) + ConvertUtil.GetValueDouble(l2[i]));
            }
            return sumList;
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
                object x;
                if (xValues == null)
                {
                    x = i + 1;
                }
                else
                {
                    x = xValues[i];
                }

                var y = yValues[i];
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
            linePath.FillColor = "none"; // No fill for line
            RenderItems.Add(linePath);
            RenderItems.AddRange(markerItems);
        }
    }

}
