using EPPlus.Export.ImageRenderer.Utils;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal abstract class ChartTypeDrawer : SvgChartObject
    {
        protected SvgChart _svgChart;
        protected ExcelChart _chartType;
        private ChartTypeDrawer(SvgChart svgChart,  ExcelChart chartType) : base(svgChart)
        {
            _svgChart = svgChart;
            _chartType = chartType;
        }
        protected List<object> LoadSeriesValues(string serieAddress, double[] numLiterals, string[] strLiterals)
        {
            List<object> values=new List<object>();
            if (numLiterals != null)
            {
                values.AddRange(numLiterals.Select(x => (object)x));
            }
            else if (strLiterals != null)
            {
                values.AddRange(strLiterals.Select(x => (object)x));
            }
            else
            {
                if(string.IsNullOrEmpty(serieAddress))
                {
                    return null;
                }
                var address = new ExcelAddressBase(serieAddress);
                var wsName = address.WorkSheetName;
                if (string.IsNullOrEmpty(wsName))
                {
                    wsName = Chart.WorkSheet.Name;
                }
                if (Chart.WorkSheet.Workbook.Worksheets[wsName] != null)
                {
                    for (int r = address.Start.Row; r <= address.End.Row; r++)
                    {
                        for (int c = address.Start.Column; c <= address.End.Column; c++)
                        {
                            values.Add(Chart.WorkSheet.Workbook.Worksheets[wsName].Cells[r, c].Value);
                        }
                    }
                }
            }
            return values; 
        }
        public List<RenderItem> RenderItems { get; } = new List<RenderItem>();
        internal static List<ChartTypeDrawer> Create(SvgChart svgChart)
        {
            var drawers = new List<ChartTypeDrawer>();
            foreach (var ct in svgChart.Chart.PlotArea.ChartTypes)
            {
                switch (ct.ChartType)
                {
                    case eChartType.Line:
                    case eChartType.LineMarkers:
                    case eChartType.LineStacked:
                    case eChartType.LineStacked100:
                    case eChartType.LineMarkersStacked:
                    case eChartType.LineMarkersStacked100:
                        drawers.Add(new LineChartTypeDrawer(svgChart, ct));
                        break;
                    default:
                        throw new NotImplementedException($"No Svg support for Chart type {ct} is implemented.");
                }
            }
            return drawers;
        }
        internal class LineChartTypeDrawer : ChartTypeDrawer
        {
            internal LineChartTypeDrawer(SvgChart svgChart, ExcelChart chartType) : base(svgChart, chartType)
            {
                foreach (ExcelLineChartSerie serie in chartType.Series)
                {
                    var yValues = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                    var xValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);
                    AddLine(chartType, serie, xValues, yValues);
                }
            }

            private void AddLine(ExcelChart chartType, ExcelLineChartSerie serie, List<object> xValues, List<object> yValues)
            {
                var xAxis= _svgChart.HorizontalAxis;
                SvgChartAxis yAxis;
                if (chartType.UseSecondaryAxis)
                {
                    yAxis = _svgChart.SecondVerticalAxis;
                }
                else
                {
                    yAxis = _svgChart.VerticalAxis;
                }
                var linePath = new SvgRenderPathItem(ChartRenderer, ChartRenderer.Bounds);
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
                    var xPos = xAxis.GetPositionInPlotarea(i);
                    var yPos = yAxis.GetPositionInPlotarea(y);

                    if(double.IsNaN(yPos)==false)
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
        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            renderItems.AddRange(RenderItems);
        }
    }
}
