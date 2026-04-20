using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;

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
                valValues = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                catValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                serCounter++;
            }

            double valTotal = 0;
            List<double> valValuesDoubles = new List<double>();

            for(int i = 0; i< valValues.Count; i++)
            {
                var origValue = valValues[i];
                var dValue = ConvertUtil.GetValueDouble(origValue, false, true);
                valValuesDoubles.Add(dValue);
                valTotal += valValuesDoubles[i];
            }

            var cx = _svgChart.Plotarea.Rectangle.Bounds.Width / 2;
            var cy = _svgChart.Plotarea.Rectangle.Bounds.Height / 2;

            var radius = Math.Min(cx, cy);

            groupItem.Bounds.Left = radius;
            groupItem.Bounds.Top = radius;

            var prevAngle = -90d;

            for (int i = 0; i < valValues.Count; i++)
            {
                valPercent.Add(valValuesDoubles[i] / valTotal);

                var angle = valPercent[i] * 360d;

                angle += prevAngle;

                circleSectorAngle.Add(angle);

                var angleRadians = angle * (Math.PI / 180.0d);

                var xPoint = cx + (radius * Math.Cos(angleRadians));
                var yPoint = cy + (radius * Math.Sin(angleRadians));
                endCoordOffsetFromCenterOfCircle.Add(new Coordinate(xPoint, yPoint));

                prevAngle = angle;
            }

            var count = Math.Min(catValues.Count, valValues.Count);

            for(int i = 0; i < chartType.Series.Count; i++)
            {
                var serie = (ExcelPieChartSerie)chartType.Series[i];
                for (var j = 0; j < catValues.Count; j++)
                {
                    var total = ConvertUtil.GetValueDouble(catValues[j], false, true);

                    var dataPoints = new List<BoundingBox>();

                    //Add the slice.
                    AddSlice(chartType, serie, catValues, valValues, dataPoints, count, j, radius);
                }
            }


            //var circ = new SvgRenderEllipseItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            //circ.Bounds.Left = cx;
            //circ.Bounds.Top = cy;

            //circ.Rx = radius;
            //circ.Ry = radius;

            //circ.Cx = cx;
            //circ.Cy = cy;

            //circ.FillColor = "purple";
            //circ.FillOpacity = 0.5d;

            //RenderItems.Add(circ);

            RenderItems.Add(new SvgEndGroupItem(ChartRenderer, null));
        }

        private void AddSlice(ExcelPieChart chartType, ExcelPieChartSerie serie, List<object> catSeries, List<object> valSeries, List<BoundingBox> dataPoints, int seriesCount, int position, double radius)
        {
            var slice = new SvgRenderPathItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            slice.BorderWidth = 2;

            //Path commands are based on percent of the Pixel height and Width
            //We must supply coords in a correct percentage of the Whole Chart

            var w = _svgChart.Plotarea.Rectangle.Bounds.Width;
            var h = _svgChart.Plotarea.Rectangle.Bounds.Height;

            var cx = (w / 2);
            var cy = (h / 2);

            //var cxWholeChart = (cx / _svgChart.Bounds.Width;

            var radX = radius / w;
            var radY = radius / h;

            var pixelWidth = _svgChart.Bounds.Width;
            var pixelHeight = _svgChart.Bounds.Height;


            var moveCenter = new PathCommands(PathCommandType.Move, slice, 0.5, 0.5);

            Coordinate startPoint;

            if(position != 0)
            {
                var lastPosX = endCoordOffsetFromCenterOfCircle[position - 1].X / w;
                var lastPosY = endCoordOffsetFromCenterOfCircle[position - 1].Y / h;
                startPoint = new Coordinate(lastPosX, lastPosY);
            }
            else
            {
                startPoint = new Coordinate(0.5, (cy - radius) / h);
            }

            var lineToStart = new PathCommands(PathCommandType.Line, slice, startPoint.X, startPoint.Y);

            var arcCommand = new PathCommands(PathCommandType.Arc, slice, new double[] { radX, radY, 0, circleSectorAngle[position] > 180 ? 1 : 0, 1, endCoordOffsetFromCenterOfCircle[position].X / w, endCoordOffsetFromCenterOfCircle[position].Y / h });

            serie.Border.Fill.Color = Color.White;
            if (position == 0)
            {
                serie.Fill.Color = Color.Red;
            }
            else if (position == 1)
            {
                serie.Fill.Color = Color.Green;
            }
            else if (position == 2)
            {
                serie.Fill.Color = Color.Blue;
            }

            slice.Commands.Add(moveCenter);
            slice.Commands.Add(lineToStart);
            slice.Commands.Add(arcCommand);

            //if(position != 0)
            //{
            //    slice.
            //}

            //slice.Commands.Add(moveCenter);
            //slice.Commands.Add(lineToEndPoint);

            slice.SetDrawingPropertiesFill(serie.Fill, chartType.StyleManager.Style.SeriesAxis.FillReference.Color);
            slice.SetDrawingPropertiesBorder(serie.Border, chartType.StyleManager.Style.SeriesAxis.BorderReference.Color, true);
            slice.SetDrawingPropertiesEffects(serie.Effect);
            RenderItems.Add(slice);
        }
    }
}
