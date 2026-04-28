using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers
{
    internal class PieChartTypeDrawer : ChartTypeDrawer
    {
        List<object> catValues = new List<object>();
        List<object> valValues = new List<object>();
        List<double> _serieValuesAsDoubles = new List<double>();
        double _totalOfSerieValues = 0;

        List<SvgPieSlice> Slices = new List<SvgPieSlice>();

        SvgGroupItemNew _groupItem;

        double _startDegrees = -90d;
        int _pieExplosionPercent = 0;

        double _sliceScaleFactor = 1.0;
        int _serCounter = 0;

        /// <summary>
        /// Radius in points
        /// </summary>
        double _radius;

        Point _circleCenter;

        public PieChartTypeDrawer(SvgChart chart, ExcelPieChart chartType) : base(chart, chartType)
        {
            _groupItem = new SvgGroupItemNew(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds.Left, _svgChart.Plotarea.Rectangle.Bounds.Top);
            //_groupItem.Position.Left = _svgChart.Plotarea.Rectangle.Bounds.Left;
            //_groupItem.Position.Top = _svgChart.Plotarea.Rectangle.Bounds.Top;
            //_groupItem.Position.Parent = _svgChart.Plotarea;

            //Read and set Starting angle offset as a rotation on the container
            //This way no rotation messes with the other calculations
            var angleOffset = double.IsNaN(chartType.FirstSliceAngle) ? 0 : chartType.FirstSliceAngle;
            _groupItem.Rotation = angleOffset;

            LoadSeriesValues(chartType);
            CalculateLocalCenterAndRadius();
            CreateIntialSlice();

            //How much to scale each slice due to pie explosion
            _sliceScaleFactor = 100d / (_pieExplosionPercent + 100d);

            var count = Math.Min(catValues.Count, valValues.Count);

            for(int i = 0; i < chartType.Series.Count; i++)
            {
                var serie = (ExcelPieChartSerie)chartType.Series[i];
                for (var j = 0; j < catValues.Count; j++)
                {
                    var total = ConvertUtil.GetValueDouble(catValues[j], false, true);

                    //Add the slice.
                    AddSlice(chartType, serie, catValues, valValues, count, j);
                }
            }

            RenderDebugEllipse();

            RenderItems.Add(_groupItem);
        }

        void RenderDebugEllipse()
        {
            var circ = new SvgRenderEllipseItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            circ.Bounds.Left = _circleCenter.Left;
            circ.Bounds.Top = _circleCenter.Top;

            circ.Rx = _radius;
            circ.Ry = _radius;

            circ.Cx = _circleCenter.Left;
            circ.Cy = _circleCenter.Top;

            circ.FillColor = "transparent";
            circ.FillOpacity = 0.3d;
            circ.BorderColor = "purple";
            circ.BorderWidth = 10;

            _groupItem.AddChildItem(circ);
        }

        Coordinate CalculateLocalPointOnCircle(double degrees)
        {
            var angleRadians = MConverter.DegreesToRadians(degrees);

            var xPoint = _circleCenter.Left + (_radius * Math.Cos(angleRadians));
            var yPoint = _circleCenter.Top + (_radius * Math.Sin(angleRadians));

            return new Coordinate(xPoint, yPoint);
        }

        void CreateIntialSlice()
        {
            //The angle of the previous slice
            //(or the 90 degree offset in the first slice)
            var prevDegrees = _startDegrees;

            for (int i = 0; i < valValues.Count; i++)
            {   //Calculate how many percent of the pie this slice is
                var valPercent = _serieValuesAsDoubles[i] / _totalOfSerieValues;
                //Create and add slice
                SvgPieSlice slice = new SvgPieSlice(ChartRenderer, _groupItem.Bounds, _circleCenter, _radius, valPercent, prevDegrees);
                Slices.Add(slice);

                //Next slice will need to be calculated starting from the degrees of this slice
                prevDegrees = slice.Degrees;
            }
        }

        void CalculateLocalCenterAndRadius()
        {
            _circleCenter = new Point();
            _circleCenter.Parent = _groupItem.Position;
            _circleCenter.Left = _svgChart.Plotarea.Rectangle.Bounds.Width / 2;
            _circleCenter.Top = _svgChart.Plotarea.Rectangle.Bounds.Height / 2;

            _groupItem.RotationPoint = _circleCenter;

            _radius = Math.Min(_circleCenter.Left, _circleCenter.Top);
        }

        void LoadSeriesValues(ExcelPieChart chartType)
        {
            //Load series values
            foreach (ExcelPieChartSerie serie in chartType.Series)
            {
                valValues = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                catValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                //Pie explosion
                _pieExplosionPercent = serie.Explosion == int.MinValue ? 0 : serie.Explosion;

                _serCounter++;
            }

            ConvertSerieValuesToDoubles();
        }

        void ConvertSerieValuesToDoubles()
        {
            for (int i = 0; i < valValues.Count; i++)
            {
                var origValue = valValues[i];
                var dValue = ConvertUtil.GetValueDouble(origValue, false, true);
                if (double.IsNaN(dValue))
                {
                    //Ignore values that are NAN. Possibly we should throw here but Excel simply seems to skip it.
                    continue;
                }
                _serieValuesAsDoubles.Add(dValue);
                _totalOfSerieValues += _serieValuesAsDoubles[i];
            }
        }

        private void AddSlice(ExcelPieChart chartType, ExcelPieChartSerie serie, List<object> catSeries, List<object> valSeries, int seriesCount, int position)
        {
            var dataPoint = serie.DataPoints[position];

            Slices[position].ImportPathData(
                _svgChart.Plotarea.Rectangle.Bounds, _svgChart.Bounds, 
                _sliceScaleFactor, dataPoint.Explosion, _pieExplosionPercent, position);

            Slices[position].ImportStlyeInfo(dataPoint, chartType);

            Slices[position].AppendGroupItem(_groupItem);
            //var innerGroup = new SvgGroupItemNew(ChartRenderer, 0, 0);

            //var slice = new SvgRenderPathItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            //slice.BorderWidth = 2;

            ////Path commands are based on percent of the Pixel height and Width
            ////We must supply coords in a correct percentage of the Whole Chart
            //var w = _svgChart.Plotarea.Rectangle.Bounds.Width;
            //var h = _svgChart.Plotarea.Rectangle.Bounds.Height;

            //var cx = (w / 2);
            //var cy = (h / 2);

            //var radX = _radius / w;
            //var radY = _radius / h;

            //var cxPercentOfTotal = (cx + _groupItem.Position.Left) / _svgChart.Bounds.Width;
            //var cyPercentOfTotal = (cy + _groupItem.Position.Top) / _svgChart.Bounds.Height;

            //var moveCenter = new PathCommands(PathCommandType.Move, slice, cxPercentOfTotal, cyPercentOfTotal);

            //Coordinate startPoint;

            //var x1 = 1d;
            //var y1 = 1d;

            //var x2 = 78d;
            //var y2 = 400d;

            //var b = Math.Pow((y2 / y1), (1 / (x2 - x1)));
            //var a = y1 / Math.Pow(b, x1);

            
            //var origin = _sliceOuterMidpoint[position];
            //innerGroup.TransformOrigin = _sliceOuterMidpoint[position];

            //innerGroup.Scale = new Coordinate(_sliceScaleFactor, _sliceScaleFactor);

            //var pointExplosion = serie.DataPoints[position].Explosion == int.MinValue ? 0 : serie.DataPoints[position].Explosion;

            ////Get directional vector
            //Graphics.Math.Vector2 pieDirection = (new Graphics.Math.Vector2(innerGroup.TransformOrigin.X, innerGroup.TransformOrigin.Y) - new Graphics.Math.Vector2((cx), (cy)));
            ////normalize the pieDirection vector
            //pieDirection = pieDirection / pieDirection.Length;

            ////If smaller than explosion direction is inward rather than outward
            //if (pointExplosion != 0 && pointExplosion < _pieExplosionPercent)
            //{
            //    pieDirection = (pieDirection * -1);
            //}
            //else if (_pieExplosionPercent != 0 && pointExplosion > _pieExplosionPercent)
            //{
            //    pointExplosion -= _pieExplosionPercent;
            //}
            ////Get distance/length to move along vector
            //Graphics.Math.Vector2 ScaledPieDirection = pieDirection * (_radius * (((double)pointExplosion / 100d)));

            //var translateX = ScaledPieDirection.X;
            //var translateY = ScaledPieDirection.Y;

            //var maxTranslationX = _svgChart.Bounds.Width - _sliceOuterMidpoint[position].X - _groupItem.Position.Left;
            ////The slices can move across whole of y 
            //var maxTranslationY = _svgChart.Bounds.Height - _sliceOuterMidpoint[position].Y;

            //var minTranslationX = -_sliceOuterMidpoint[position].X -_groupItem.Position.Left;
            //var minTranslationY = -_sliceOuterMidpoint[position].Y - _groupItem.Position.Top;

            //Graphics.Point lengthPoint = new Graphics.Point();

            //if(ScaledPieDirection.X != 0)
            //{
            //    if (ScaledPieDirection.X > 0 && ScaledPieDirection.X > maxTranslationX)
            //    {
            //        lengthPoint.Left = maxTranslationX;
            //    }
            //    else if (ScaledPieDirection.X < minTranslationX)
            //    {
            //        lengthPoint.Left = Math.Abs(minTranslationX);
            //    }
            //}

            //if (ScaledPieDirection.Y != 0)
            //{
            //    if (ScaledPieDirection.Y > 0 && ScaledPieDirection.Y > maxTranslationY)
            //    {
            //        lengthPoint.Top = maxTranslationY;
            //    }
            //    else if (ScaledPieDirection.Y < minTranslationY)
            //    {
            //        lengthPoint.Top = Math.Abs(minTranslationY);
            //    }
            //}

            //var smallestLength = Math.Min(lengthPoint.Left, lengthPoint.Top);
            //if(smallestLength == 0)
            //{
            //    smallestLength = Math.Max(lengthPoint.Left, lengthPoint.Top);
            //}

            //double translationLeft;
            //double translationTop;

            //if (smallestLength != 0 && smallestLength < ScaledPieDirection.Length)
            //{
            //    var normalizedVector = ScaledPieDirection / ScaledPieDirection.Length;
            //    var appliedVector = normalizedVector * smallestLength;

            //    translationLeft = appliedVector.X;
            //    translationTop = appliedVector.Y;
            //}
            //else
            //{
            //    translationLeft = Math.Min(translateX, maxTranslationX);
            //    translationTop = Math.Min(translateY, maxTranslationY);

            //    translationLeft = Math.Max(translationLeft, minTranslationX);
            //    translationTop = Math.Max(translationTop, minTranslationY);
            //}

            //innerGroup.Position.Left = translationLeft;
            //innerGroup.Position.Top = translationTop;

            //if (position != 0)
            //{
            //    var lastPosX = _endCoordOffsetFromLocalCenterOfCircle[position - 1].X / w;
            //    var lastPosY = _endCoordOffsetFromLocalCenterOfCircle[position - 1].Y / h;
            //    startPoint = new Coordinate(lastPosX, lastPosY);
            //}
            //else
            //{
            //    startPoint = new Coordinate(cxPercentOfTotal, (cy - _radius) / h);
            //}

            //var lineToStart = new PathCommands(PathCommandType.Line, slice, startPoint.X, startPoint.Y);

            //var individualAngle = valPercent[position] * 360d;

            //var arcCommand = new PathCommands(PathCommandType.Arc, slice, new double[] { radX, radY, 0, individualAngle > 180 ? 1 : 0, 1, _endCoordOffsetFromLocalCenterOfCircle[position].X / w, _endCoordOffsetFromLocalCenterOfCircle[position].Y / h });
            //var end = new PathCommands(PathCommandType.End, slice, _endCoordOffsetFromLocalCenterOfCircle[position].X / w, _endCoordOffsetFromLocalCenterOfCircle[position].Y / h);

            //slice.Commands.Add(moveCenter);
            //slice.Commands.Add(lineToStart);
            //slice.Commands.Add(arcCommand);
            //slice.Commands.Add(end);

            //slice.SetDrawingPropertiesFill(serie.DataPoints[position].Fill, chartType.StyleManager.Style.DataPoint.FillReference.Color);
            //slice.SetDrawingPropertiesBorder(serie.DataPoints[position].Border, chartType.StyleManager.Style.DataPoint.BorderReference.Color, true);
            //slice.SetDrawingPropertiesEffects(serie.DataPoints[position].Effect);
            //innerGroup.AddChildItem(slice);

            //_groupItem.AddChildItem(innerGroup);


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
