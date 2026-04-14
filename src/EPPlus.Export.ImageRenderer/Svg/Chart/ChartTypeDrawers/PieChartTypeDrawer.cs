using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers
{
    internal class PieChartTypeDrawer : ChartTypeDrawer
    {
        List<object> catValues = new List<object>();
        List<object> valValues = new List<object>();
        List<double> valPercent = new List<double>();
        List<double> circleSectorAngle = new List<double>();

        List<Coordinate> endCoordOffsetFromCenterOfCircle = new List<Coordinate>();

        public PieChartTypeDrawer(SvgChart chart, ExcelPieChart chartType) : base(chart, chartType)
        {
            var groupItem = new SvgGroupItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);
            RenderItems.Add(groupItem);
            var xValues = new List<List<object>>();
            var yValues = new List<List<object>>();
            int serCounter = 0;

            foreach (ExcelPieChartSerie serie in chartType.Series)
            {
                List<object> valValue, catValue;

                valValue = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                catValue = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                catValues.Add(catValue);
                valValues.Add(valValue);

                serCounter++;
            }

            double valTotal = 0;
            List<double> valValuesDoubles = new List<double>();

            for(int i = 0; i< valValues.Count; i++)
            {
                valValuesDoubles.Add(ConvertUtil.GetValueDouble(valValues[i], false, true));
                valTotal += valValuesDoubles[i];
            }

            var cx = _svgChart.Plotarea.Rectangle.Bounds.Width / 2;
            var cy = _svgChart.Plotarea.Rectangle.Bounds.Height / 2;

            var radius = Math.Min(_svgChart.Plotarea.Rectangle.Bounds.Height, _svgChart.Plotarea.Rectangle.Bounds.Width);


            for (int i = 0; i < valValues.Count; i++)
            {
                valPercent.Add(valValuesDoubles[i] / valTotal);
                circleSectorAngle.Add(valPercent[i] / 360d);
                var xPoint = cx + (radius * Math.Cos(circleSectorAngle[i]));
                var yPoint = cy + (radius * Math.Sin(circleSectorAngle[i]));
                endCoordOffsetFromCenterOfCircle.Add(new Coordinate(xPoint, yPoint));
            }

            var count = Math.Min(catValues.Count, valValues.Count);
            for (var i = 0; i < catValues.Count; i++)
            {
                var serie = (ExcelPieChartSerie)chartType.Series[i];

                var total = ConvertUtil.GetValueDouble(catValues[i], false, true);

                var dataPoints = new List<BoundingBox>();

                //Add the slice.
                AddSlice(chartType, serie, catValues, valValues, dataPoints, count, i, radius);


            }
            RenderItems.Add(new SvgEndGroupItem(ChartRenderer, null));
        }

        private void AddSlice(ExcelPieChart chartType, ExcelPieChartSerie serie, List<object> catSeries, List<object> valSeries, List<BoundingBox> dataPoints, int seriesCount, int position, double radius)
        {
            var slice = new SvgRenderEllipseItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);


            slice.SetDrawingPropertiesFill(serie.Fill, chartType.StyleManager.Style.SeriesAxis.FillReference.Color);
            slice.SetDrawingPropertiesBorder(serie.Border, chartType.StyleManager.Style.SeriesAxis.BorderReference.Color, true);
            slice.SetDrawingPropertiesEffects(serie.Effect);
            RenderItems.Add(slice);
        }
    }
}
