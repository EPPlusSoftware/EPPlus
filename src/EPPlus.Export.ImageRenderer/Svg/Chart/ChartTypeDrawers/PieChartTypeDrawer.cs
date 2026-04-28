using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
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
using System.IO.Pipes;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers
{
    internal class PieChartTypeDrawer : ChartTypeDrawer
    {
        List<object> catValues = new List<object>();
        List<object> valValues = new List<object>();
        List<double> valPercent = new List<double>();
        List<double> circleSectorAngle = new List<double>();

        SvgGroupItemNew groupItem;

        List<Coordinate> endCoordOffsetFromCenterOfCircle = new List<Coordinate>();
        List<Coordinate> sliceTransformOrigin = new List<Coordinate>();

        double _startAngle;

        double _sliceScaleFactor = 1.0;

        public PieChartTypeDrawer(SvgChart chart, ExcelPieChart chartType) : base(chart, chartType)
        {
            groupItem = new SvgGroupItemNew(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds.Left, _svgChart.Plotarea.Rectangle.Bounds.Top);

            var xValues = new List<List<object>>();
            var yValues = new List<List<object>>();
            int serCounter = 0;

            var angleOffset = double.IsNaN(chartType.FirstSliceAngle) ? 0 : chartType.FirstSliceAngle;

            groupItem.Rotation = angleOffset;

            _startAngle = -90d;
            int pieExplosionPercent = 0;


            foreach (ExcelPieChartSerie serie in chartType.Series)
            {
                valValues = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                catValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                //Pie explosion
                pieExplosionPercent = serie.Explosion == int.MinValue ? 0 : serie.Explosion;

                serCounter++;
            }

            double valTotal = 0;
            List<double> valValuesDoubles = new List<double>();

            for(int i = 0; i< valValues.Count; i++)
            {
                var origValue = valValues[i];
                var dValue = ConvertUtil.GetValueDouble(origValue, false, true);
                if(double.IsNaN(dValue))
                {
                    //Ignore values that are NAN. Possibly we should throw here but Excel simply seems to skip it.
                    continue;
                }
                valValuesDoubles.Add(dValue);
                valTotal += valValuesDoubles[i];
            }

            var cx = (_svgChart.Plotarea.Rectangle.Bounds.Width / 2);
            var cy = (_svgChart.Plotarea.Rectangle.Bounds.Height / 2);

            groupItem.RotationPoint = new Graphics.Point(cx, cy);

            var radius = Math.Min(cx, cy);

            var prevAngle = _startAngle;

            for (int i = 0; i < valValues.Count; i++)
            {
                valPercent.Add(valValuesDoubles[i] / valTotal);

                var angle = valPercent[i] * 360d;
                var halfAngle = angle / 2;

                angle += prevAngle;

                circleSectorAngle.Add(angle);

                var angleRadians = angle * (Math.PI / 180.0d);

                var xPoint = cx + (radius * Math.Cos(angleRadians));
                var yPoint = cy + (radius * Math.Sin(angleRadians));
                endCoordOffsetFromCenterOfCircle.Add(new Coordinate(xPoint, yPoint));

                //Add prev at this point since we don't want to halve prev angle
                halfAngle += prevAngle;

                var halfAngleRadians = halfAngle * (Math.PI / 180.0d);

                var xPointHalf = cx + (radius * Math.Cos(halfAngleRadians));
                var yPointHalf = cy + (radius * Math.Sin(halfAngleRadians));

                sliceTransformOrigin.Add(new Coordinate(xPointHalf, yPointHalf));

                prevAngle = angle;
            }

            _sliceScaleFactor = 100d / (pieExplosionPercent + 100d);

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

            //circ.FillColor = "transparent";
            //circ.FillOpacity = 0.3d;
            //circ.BorderColor = "purple";
            //circ.BorderWidth = 10;

            //groupItem.AddChildItem(circ);
            RenderItems.Add(groupItem);
        }

        private void AddSlice(ExcelPieChart chartType, ExcelPieChartSerie serie, List<object> catSeries, List<object> valSeries, List<BoundingBox> dataPoints, int seriesCount, int position, double radius)
        {
            var innerGroup = new SvgGroupItemNew(ChartRenderer, 0, 0);

            var slice = new SvgRenderPathItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            slice.BorderWidth = 2;

            //Path commands are based on percent of the Pixel height and Width
            //We must supply coords in a correct percentage of the Whole Chart

            var w = _svgChart.Plotarea.Rectangle.Bounds.Width;
            var h = _svgChart.Plotarea.Rectangle.Bounds.Height;

            var cx = (w / 2);
            var cy = (h / 2);

            var radX = radius / w;
            var radY = radius / h;

            var cxPercentOfTotal = (cx + groupItem.Position.Left) / _svgChart.Bounds.Width;
            var cyPercentOfTotal = (cy + groupItem.Position.Top) / _svgChart.Bounds.Height;

            var moveCenter = new PathCommands(PathCommandType.Move, slice, cxPercentOfTotal, cyPercentOfTotal);

            Coordinate startPoint;

            var x1 = 1d;
            var y1 = 1d;

            var x2 = 78d;
            var y2 = 400d;

            var b = Math.Pow((y2 / y1), (1 / (x2 - x1)));
            var a = y1 / Math.Pow(b, x1);

            
            var origin = sliceTransformOrigin[position];
            innerGroup.TransformOrigin = sliceTransformOrigin[position];

            innerGroup.Scale = new Coordinate(_sliceScaleFactor, _sliceScaleFactor);

            var pointExplosion = serie.DataPoints[position].Explosion == int.MinValue ? 0 : serie.DataPoints[position].Explosion;

            //Get directional vector
            Graphics.Math.Vector2 pieDirection = (new Graphics.Math.Vector2(innerGroup.TransformOrigin.X, innerGroup.TransformOrigin.Y) - new Graphics.Math.Vector2((cx), (cy)));
            //Get distance/length to move along vector
            Graphics.Math.Vector2 ScaledPieDirection = pieDirection * (((double)pointExplosion / 100d));

            Graphics.Point transformPoint = new Graphics.Point(innerGroup.TransformOrigin.X, innerGroup.TransformOrigin.Y);

            var translateX = ScaledPieDirection.X;
            var translateY = ScaledPieDirection.Y;

            var maxTranslationX = _svgChart.Bounds.Width - sliceTransformOrigin[position].X;
            var maxTranslationY = _svgChart.Bounds.Height - sliceTransformOrigin[position].Y;

            var minTranslationX = -sliceTransformOrigin[position].X;
            var minTranslationY = -sliceTransformOrigin[position].Y;

            Graphics.Point lengthPoint = new Graphics.Point();

            if(ScaledPieDirection.X != 0)
            {
                if (ScaledPieDirection.X > 0 && ScaledPieDirection.X > maxTranslationX)
                {
                    lengthPoint.Left = maxTranslationX - sliceTransformOrigin[position].X;
                }
                else if (ScaledPieDirection.X < minTranslationX)
                {
                    lengthPoint.Left = Math.Abs(sliceTransformOrigin[position].X + minTranslationX);
                }
            }

            if (ScaledPieDirection.Y != 0)
            {
                if (ScaledPieDirection.Y > 0 && ScaledPieDirection.Y > maxTranslationY)
                {
                    lengthPoint.Top = maxTranslationY - sliceTransformOrigin[position].Y;
                }
                else if (ScaledPieDirection.Y < minTranslationY)
                {
                    lengthPoint.Top = Math.Abs(sliceTransformOrigin[position].Y + minTranslationY);
                }
            }

            var smallestLength = Math.Min(lengthPoint.Left, lengthPoint.Top);

            //if(smallestLength != 0)
            //{
            //    transformPoint.Scale = 
            //}

            var translationLeft = Math.Min(translateX, maxTranslationX);
            var translationTop = Math.Min(translateY, maxTranslationY);

            translationLeft = Math.Max(translationLeft, minTranslationX);
            translationTop = Math.Max(translationTop, minTranslationY);

            innerGroup.Position.Left = translationLeft;
            innerGroup.Position.Top = translationTop;

            if (position != 0)
            {
                var lastPosX = endCoordOffsetFromCenterOfCircle[position - 1].X / w;
                var lastPosY = endCoordOffsetFromCenterOfCircle[position - 1].Y / h;
                startPoint = new Coordinate(lastPosX, lastPosY);
            }
            else
            {
                startPoint = new Coordinate(cxPercentOfTotal, (cy - radius) / h);
            }

            var lineToStart = new PathCommands(PathCommandType.Line, slice, startPoint.X, startPoint.Y);

            var individualAngle = valPercent[position] * 360d;

            var arcCommand = new PathCommands(PathCommandType.Arc, slice, new double[] { radX, radY, 0, individualAngle > 180 ? 1 : 0, 1, endCoordOffsetFromCenterOfCircle[position].X / w, endCoordOffsetFromCenterOfCircle[position].Y / h });
            var end = new PathCommands(PathCommandType.End, slice, endCoordOffsetFromCenterOfCircle[position].X / w, endCoordOffsetFromCenterOfCircle[position].Y / h);

            slice.Commands.Add(moveCenter);
            slice.Commands.Add(lineToStart);
            slice.Commands.Add(arcCommand);
            slice.Commands.Add(end);

            slice.SetDrawingPropertiesFill(serie.DataPoints[position].Fill, chartType.StyleManager.Style.DataPoint.FillReference.Color);
            slice.SetDrawingPropertiesBorder(serie.DataPoints[position].Border, chartType.StyleManager.Style.DataPoint.BorderReference.Color, true);
            slice.SetDrawingPropertiesEffects(serie.DataPoints[position].Effect);
            innerGroup.AddChildItem(slice);

            groupItem.AddChildItem(innerGroup);


            //var LineItem = new SvgRenderPathItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            //var midPointLine = new PathCommands(PathCommandType.Line, LineItem, origin.X / w, origin.Y / h);
            //var moveCenterLine = new PathCommands(PathCommandType.Move, LineItem, cxPercentOfTotal, cyPercentOfTotal);
            //LineItem.Commands.Add(moveCenterLine);
            //LineItem.Commands.Add(midPointLine);

            //LineItem.BorderColor = "purple";
            //LineItem.BorderWidth = 3;

            //groupItem.AddChildItem(LineItem);
        }
    }
}
