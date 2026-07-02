using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils.Drawing;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class PieSliceRenderItem : ChartDrawingObject
    {
        double _radius;
        /// <summary>
        /// The holder of transform-origin/translations
        /// </summary>
        GroupRenderItem _innerGroup;
        /// <summary>
        /// The holder of the actual items, AFTER origin/translations
        /// </summary>
        GroupRenderItem _innerItems;
        Point _circleCenter;

        /// <summary>
        /// How many percent of the pie this represents
        /// </summary>
        double _percent;

        internal double Degrees { get; private set; }

        Point _startPoint;
        Point _startPointHalf;
        Point _midPoint;
        Point _endPoint;
        Point _endPointHalf;

        internal override System.Drawing.Color? DefaultFillColor { get; }

        /// <summary>
        /// Get copy of start point of slice in local coordinates
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetStartPointPositionLocal()
        {
            return new Coordinate(_startPoint.Left, _startPoint.Top);
        }
        /// <summary>
        /// Get copy of mid point of slice in local coordinates
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetMidPointLocal()
        {
            return new Coordinate(_midPoint.Left, _midPoint.Top);
        }
        /// <summary>
        /// Get copy of end point of slice in local coordinates
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetEndPointLocal()
        {
            return new Coordinate(_endPoint.LocalPosition.X, _endPoint.LocalPosition.Y);
        }

        PathRenderItem _slicePath;


        List<RenderItem> DebugItems;
        PathRenderItem _debugBoundsPath;
        RectRenderItem _debugCircleCenter;

        bool ExistWithinRange(double target, double min, double max)
        {
            if (min < target && target < max)
            {
                return true;
            }
            return false;
        }

        internal BoundingBox ExtremePoints { get; set; }

        void CalculateWidthHeight(double prevSliceDegrees)
        {

            var endPointDegrees = prevSliceDegrees + Degrees;
            if (endPointDegrees < 0)
            {
                endPointDegrees = 360 + endPointDegrees;
            }

            var startPointDegrees = prevSliceDegrees;
            if (startPointDegrees < 0)
            {
                startPointDegrees = 360 + startPointDegrees;
            }

            var circleSectorDegrees = endPointDegrees - startPointDegrees;

            double maxX;
            double maxY;
            double minY;
            double minX;

            if (ExistWithinRange(90, startPointDegrees, endPointDegrees))
            {
                maxY = _circleCenter.Top + _radius;
            }
            else
            {
                maxY = Math.Max(_startPoint.Top, _endPoint.Top);
            }

            maxY = Math.Max(maxY, _circleCenter.Top);

            if (ExistWithinRange(180, startPointDegrees, endPointDegrees))
            {
                minX = _circleCenter.Left - _radius;
            }
            else
            {
                minX = Math.Min(_startPoint.Left, _endPoint.Left);
            }

            minX = Math.Min(minX, _circleCenter.Left);

            if (ExistWithinRange(270, startPointDegrees, endPointDegrees))
            {
                minY = _circleCenter.Top - _radius;
            }
            else
            {
                minY = Math.Min(_startPoint.Top, _endPoint.Top);
            }

            minY = Math.Min(minY, _circleCenter.Top);

            if (endPointDegrees < startPointDegrees || ExistWithinRange(0, startPointDegrees, endPointDegrees))
            {
                if (endPointDegrees > 270)
                {
                    maxX = _circleCenter.Left;
                }
                else
                {
                    maxX = _circleCenter.Left + _radius;
                }
            }
            else
            {
                maxX = Math.Max(_startPoint.Left, _endPoint.Left);
            }

            maxX = Math.Max(_circleCenter.Left, maxX);


            ExtremePoints = new BoundingBox(minX, minY, maxX - minX, maxY - minY);
            ExtremePoints.Parent = _innerGroup.Bounds;
        }

        private double _sliceScaleFactor = 1d;
        private double _scaledRadius { get { return _radius * _sliceScaleFactor; } }


        private void CalculateExplosionDir()
        {
            var transformOriginLocal = new Vector2(_innerGroup.TransformOrigin.X, _innerGroup.TransformOrigin.Y);

            //Get directional vector (in local coords but does not matter since we make it directional)
            Vector2 pieDirection = transformOriginLocal - _circleCenter.LocalPosition;

            //normalize the pieDirection vector so that it is percentual and with lenght == 1
            pieDirection = pieDirection / pieDirection.Length;

            CtrToOuterMidDir = new Vector2(pieDirection.X, pieDirection.Y);
        }

        public PieSliceRenderItem(ChartRenderer renderer, BoundingBox parent, Point circleCenter, double radius, double percentOfPie, double prevSliceDegrees) : base(renderer)
        {
            DefaultFillColor = renderer.Theme.ColorScheme.Accent1.GetColor();
            Rectangle.Bounds.Parent = parent;
            _radius = radius;
            _percent = percentOfPie;
            //How many degrees that percentage is out of 360
            Degrees = _percent * 360d;

            _innerGroup = new GroupRenderItem(parent, 0, circleCenter);
            _innerGroup.Bounds.Parent = _innerGroup.TranslationOffset;
            _innerGroup.Bounds.Name = "InnerGroupChartDrawer";

            _circleCenter = circleCenter;

            _startPoint = CalculateLocalPointOnCircle(prevSliceDegrees);
            _startPointHalf = CalculateLocalPointOnCircleHalfRadius(prevSliceDegrees);

            //The degrees of the midpoint
            var halfDegrees = Degrees / 2;

            _endPoint = CalculateLocalPointOnCircle(Degrees + prevSliceDegrees);
            _endPointHalf = CalculateLocalPointOnCircleHalfRadius(Degrees + prevSliceDegrees);

            //We add prev at this point since we don't want to halve the previous angle only the current one
            _midPoint = CalculateLocalPointOnCircle(halfDegrees + prevSliceDegrees);

            //We must calculate transforms from the outer midpoint.
            //This is to ensure that point never leaves the parent container
            _innerGroup.TransformOrigin = GetMidPointLocal();

            CalculateExplosionDir();
            CalculateWidthHeight(prevSliceDegrees);


            _innerItems = new GroupRenderItem(_innerGroup.Bounds, 0);
        }

        internal void ImportPathData(BoundingBox plotAreaBounds, BoundingBox globalAreaBounds, double sliceScaleFactor, double explosionOfPoint, double pieExplosion, int position)
        {
            _slicePath = new PathRenderItem(plotAreaBounds);

            _slicePath.BorderWidth = 5;

            //Calculate path commands
            var moveCenter = new PathCommands(PathCommandType.Move, _circleCenter.Left, _circleCenter.Top);
            var lineToStart = new PathCommands(PathCommandType.Line, _startPoint.Left, _startPoint.Top);

            var arcCommand = new PathCommands(PathCommandType.Arc, new double[] { _radius, _radius, 0, Degrees > 180 ? 1 : 0, 1, _endPoint.Left, _endPoint.Top });
            var end = new PathCommands(PathCommandType.End, _endPoint.Left, _endPoint.Top);

            //Get max and min values
            var localMax = GetTranslationMaxLocal(globalAreaBounds.Width, globalAreaBounds.Height);
            var localMin = GetTranslationMinLocal(globalAreaBounds.Width, globalAreaBounds.Height);

            _sliceScaleFactor = sliceScaleFactor;
            //Translate and scale path
            _innerGroup.Scale = new Coordinate(_sliceScaleFactor, _sliceScaleFactor);
            CalculatePointExplosion(explosionOfPoint, pieExplosion, localMax, localMin);
            CalculateLargestRectWithinCircleSegment();

            //Add the actual commands
            _slicePath.Commands.Add(moveCenter);
            _slicePath.Commands.Add(lineToStart);
            _slicePath.Commands.Add(arcCommand);

            //Change to != -1 to activate debug items
            if (position == -1)
            {
                //Visualize all points
                AddDebugLines(moveCenter, plotAreaBounds);
            }

            _slicePath.Commands.Add(end);

        }

        /// <summary>
        /// Adds line from center to outer mid point (transform-origin) and to end point
        /// AKA line along scale/explosion vector
        /// </summary>
        /// <param name="moveCenter"></param>
        private void AddDebugLines(PathCommands moveCenter, BoundingBox bounds)
        {
            DebugItems = new List<RenderItem>();

            //Render bounds for slice
            _debugBoundsPath = new PathRenderItem(bounds);
            _debugBoundsPath.BorderColor = "red";
            _debugBoundsPath.FillColor = "transparent";
            _debugBoundsPath.BorderWidth = 3;
            var moveCenterDebug = new PathCommands(PathCommandType.Move, ExtremePoints.Left, ExtremePoints.Top);
            _debugBoundsPath.Commands.Add(moveCenterDebug);

            //Draw extremes/bounds
            //var lineToTopLeft = new PathCommands(PathCommandType.Line, ExtremePoints.Left, ExtremePoints.Top);
            var lineToTopRight = new PathCommands(PathCommandType.Line, ExtremePoints.Right, ExtremePoints.Top);

            var lineToBottomRight = new PathCommands(PathCommandType.Line, ExtremePoints.Right, ExtremePoints.Bottom);
            var lineToBottomLeft = new PathCommands(PathCommandType.Line, ExtremePoints.Left, ExtremePoints.Bottom);
            var end = new PathCommands(PathCommandType.End, ExtremePoints.Left, ExtremePoints.Top);

            _debugBoundsPath.Commands.Add(lineToTopRight);
            _debugBoundsPath.Commands.Add(lineToBottomRight);
            _debugBoundsPath.Commands.Add(lineToBottomLeft);
            _debugBoundsPath.Commands.Add(end);

            DebugItems.Add(_debugBoundsPath);

            //Render green dot at circle center
            _debugCircleCenter = GenerateDebugPoint(bounds, new Coordinate(_circleCenter.Left, _circleCenter.Top), "green");

            DebugItems.Add(_debugCircleCenter);

            //render dots at points
            var pointColor = "yellow";
            
            var startDebug = GenerateDebugPoint(bounds, GetStartPointPositionLocal(), pointColor);
            var midDebug = GenerateDebugPoint(bounds, GetMidPointLocal(), pointColor);
            var endDebug = GenerateDebugPoint(bounds, GetEndPointLocal(), pointColor);

            DebugItems.Add(startDebug);
            DebugItems.Add(midDebug);
            DebugItems.Add(endDebug);
        }

        private RectRenderItem GenerateDebugPoint(BoundingBox parent, Coordinate point, string fillColor)
        {
            double l = -2.5d;
            double t = -2.5d;
            double w = 5d;
            double h = 5d;

            return new RectRenderItem(parent) { Left = l + point.X, Top = t + point.Y, Width = w, Height = h, FillColor = fillColor };
        }


        internal void ImportStlyeInfo(ExcelPieChartSerie serie, ExcelPieChart chartType, int position)
        {

            var defaultFill = DefaultFillColor;

            if (position >= 0 && serie.DataPoints.ContainsKey(position))
            {
                var dp = serie.DataPoints[position];
                ChartTypeDrawer.SetFillDataPoint(Chart, serie, position, _slicePath, dp, Chart.StyleManager.Style?.SeriesLine);
            }
            else
            {
                ChartTypeDrawer.SetFillSerie(Chart, chartType, serie, 0, position, _slicePath);
            }
            //if(chartType.VaryColors)
            //{
            //    if(chartType.StyleManager.Style == null)
            //    {
            //        var mod5 = position % 5;
            //        //TODO: Only works for base-case. Add support for patterns 1,3 and 4 instead of just 2 as basecase
            //        defaultFill = ChartRenderer.Theme.ColorScheme.GetColorByEnum(OfficeOpenXml.Drawing.eSchemeColor.Accent1 + mod5).GetColor();
            //    }
            //}
            //_slicePath.SetDrawingPropertiesFill(ChartRenderer.Theme, dp.Fill, chartType.StyleManager.Style?.DataPoint.FillReference.Color, false, defaultFill);
            //_slicePath.SetDrawingPropertiesBorder(ChartRenderer.Theme, dp.Border, chartType.StyleManager.Style?.DataPoint.BorderReference.Color, true);
            //_slicePath.SetDrawingPropertiesEffects(ChartRenderer.Theme, dp.Effect);
        }

        internal void AppendGroupItem(GroupRenderItem group)
        {
            //Apply translation after all calculations are done
            _innerGroup.Left += _innerGroup.TranslationOffset.Left;
            _innerGroup.Top += _innerGroup.TranslationOffset.Top;

            //The slice items post transform operations
            _innerItems.AddChildItem(_slicePath);
            //The bounds and translations of the slice
            _innerGroup.AddChildItem(_innerItems);

            if (DebugItems != null && DebugItems.Count > 0)
            {
                foreach(var debugItem in DebugItems)
                {
                    _innerGroup.AddChildItem(debugItem);
                }
            }

            //The group containing all slices
            group.AddChildItem(_innerGroup);
        }

        Vector2 GetTranslationMaxLocal(double globalWidth, double globalHeight)
        {
            var worldPositionTransformOrigin = _midPoint.Position;

            //Calculate extremes 
            Point localMax = new Point(
                globalWidth - worldPositionTransformOrigin.X,
                globalHeight - worldPositionTransformOrigin.Y);

            return localMax.LocalPosition;
        }

        Vector2 GetTranslationMinLocal(double globalWidth, double globalHeight)
        {
            var worldPositionTransformOrigin = _midPoint.Position;
            //Calculate extremes 
            Point worldMin = new Point(-worldPositionTransformOrigin.X, -worldPositionTransformOrigin.Y);

            //var localMin = _innerGroup.Position.Parent.TransformPointToLocal(worldMin.Position);
            return worldMin.LocalPosition;
        }

        /// <summary>
        /// The directional vector from center of circle to the outer midpoint of the pie slice
        /// </summary>
        internal Vector2 CtrToOuterMidDir { get; private set; }

        /// <summary>
        /// Gets the whole vector with length to the end
        /// </summary>
        /// <returns></returns>
        internal Vector2 GetWholeVectorCenterToMid()
        {
            var translationVector = GetLocalTranslationVector(100);
            var pt = translationVector;
            return pt;
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        internal Point GetSliceShapeCenterLocal()
        {
            var translationVector = GetLocalTranslationVector(50);
            var pt = _circleCenter.LocalPosition + translationVector;
            var SliceCenterLocal = new Point(pt.X, pt.Y);

            return SliceCenterLocal;
        }
        /// <summary>
        /// Input must be between 0 and 100
        /// </summary>
        /// <param name="percentTowardsEndPoint"></param>
        /// <returns></returns>
        Vector2 GetLocalTranslationVector(double percentTowardsEndPoint)
        {
            if (percentTowardsEndPoint < 0 || percentTowardsEndPoint > 100)
            {
                throw new InvalidOperationException($"input: '{percentTowardsEndPoint}' invalid. Must be between 0 and 100");
            }

            var moveFactor = (double)percentTowardsEndPoint / 100d;
            var moveFactoredDirection = moveFactor * CtrToOuterMidDir;

            //Get distance/length to move along vector. We translate according to the scaled down radius
            Vector2 LocalTranslationVector = moveFactoredDirection * _scaledRadius;
            return LocalTranslationVector;
        }

        Vector2 GetLocalTranslationVector(double explosionOfPoint, double pieExplosion)
        {
            //Get point explosion value
            var pointExplosion = explosionOfPoint == int.MinValue ? 0 : explosionOfPoint;

            var pieDirection = new Vector2(CtrToOuterMidDir.X, CtrToOuterMidDir.Y);

            if (pointExplosion != 0 && pointExplosion < pieExplosion)
            {
                //Direction is inward rather than outward
                pieDirection = (pieDirection * -1);
            }
            else if (pieExplosion != 0 && pointExplosion > pieExplosion)
            {
                //Scaling has already translated the slice partially.
                //Remove that from the translation percent
                pointExplosion -= pieExplosion;
            }

            var ptExplodeFactor = (double)pointExplosion / 100d;
            var ptFactoredDirection = ptExplodeFactor * pieDirection;

            //Get distance/length to move along vector. We translate according to the scaled down radius
            Vector2 LocalTranslationVector = ptFactoredDirection * _scaledRadius;
            return LocalTranslationVector;
        }

        Coordinate GetFinalLocalTranslation(Vector2 LocalTranslationVector, Vector2 localMax, Vector2 localMin)
        {
            Coordinate lengthPoint = new Coordinate(0, 0);

            //Check if local is above or below extremes in X axis
            if (LocalTranslationVector.X != 0)
            {
                if (LocalTranslationVector.X > 0 && LocalTranslationVector.X > localMax.X)
                {
                    lengthPoint.X = localMax.X;
                }
                else if (LocalTranslationVector.X < localMin.X)
                {
                    lengthPoint.X = Math.Abs(localMin.X);
                }
            }

            //Check if local is above or below extremes in Y axis
            if (LocalTranslationVector.Y != 0)
            {
                if (LocalTranslationVector.Y > 0 && LocalTranslationVector.Y > localMax.Y)
                {
                    lengthPoint.Y = localMax.Y;
                }
                else if (LocalTranslationVector.Y < localMin.Y)
                {
                    lengthPoint.Y = Math.Abs(localMin.Y);
                }
            }

            //Find the smallest length of a vector that goes beyond the extremes
            //In case both do
            var maxAllowedLength = Math.Min(lengthPoint.X, lengthPoint.Y);
            if (maxAllowedLength == 0)
            {
                //Avoid issues if one axis is 0 and the vector that goes over is positive
                maxAllowedLength = Math.Max(lengthPoint.X, lengthPoint.Y);
            }

            double translationLeft;
            double translationTop;

            if (maxAllowedLength != 0 && maxAllowedLength < LocalTranslationVector.Length)
            {
                //If the length is larger than maximum allowed length we stop applying translation
                var normalizedVector = LocalTranslationVector / LocalTranslationVector.Length;
                var appliedVector = normalizedVector * maxAllowedLength;

                translationLeft = appliedVector.X;
                translationTop = appliedVector.Y;
            }
            else
            {
                //The length is within the bounds. No binding neccesary.
                translationLeft = LocalTranslationVector.X;
                translationTop = LocalTranslationVector.Y;
            }

            return new Coordinate(translationLeft, translationTop);
        }

        void CalculatePointExplosion(double explosionOfPoint, double pieExplosion, Vector2 localMax,Vector2 localMin)
        {
            //Get distance/length to move along vector
            Vector2 LocalTranslationVector = GetLocalTranslationVector(explosionOfPoint, pieExplosion);
            var finalTranslation = GetFinalLocalTranslation(LocalTranslationVector, localMax, localMin);

            _innerGroup.TranslationOffset.Left = finalTranslation.X;
            _innerGroup.TranslationOffset.Top = finalTranslation.Y;
        }

        internal double LargestWidthRectangle { get; private set; }
        internal double LargestHeightRectangle { get; private set; }

        void CalculateLargestRectWithinCircleSegment()
        {
            //Calculate thetha = alpha/4
            var angleForTriangle = Degrees / 4d;

            var heightTriangle = Math.Sin(MConverter.DegreesToRadians(angleForTriangle)) * _radius + 1; // add 1 for small rounding fault making too small
            LargestWidthRectangle = Math.Cos(MConverter.DegreesToRadians(angleForTriangle)) * _radius;
            LargestHeightRectangle = heightTriangle * 2;
        }

        Point CalculateLocalPointOnCircle(double degrees)
        {
            var angleRadians = MConverter.DegreesToRadians(degrees);

            var xPoint = _circleCenter.Left + (_radius * Math.Cos(angleRadians));
            var yPoint = _circleCenter.Top + (_radius * Math.Sin(angleRadians));

            var point = new Point();

            //Ensure the cx/cy offset
            point.Parent = _innerGroup.TranslationOffset.Parent;
            point.Left = xPoint;
            point.Top = yPoint;

            return point;
        }

        Point CalculateLocalPointOnCircleHalfRadius(double degrees)
        {
            var angleRadians = MConverter.DegreesToRadians(degrees);

            var xPoint = _circleCenter.Left + (_radius/2d * Math.Cos(angleRadians));
            var yPoint = _circleCenter.Top + (_radius/2d * Math.Sin(angleRadians));

            var point = new Point();

            //Ensure the cx/cy offset
            point.Parent = _innerGroup.TranslationOffset.Parent;
            point.Left = xPoint;
            point.Top = yPoint;

            return point;
        }

        internal Coordinate GetOuterMidpointInGlobalCoords()
        {
            return new Coordinate(_midPoint.Position.X, _midPoint.Position.Y);
        }


        internal Transform CopyOuterMidPoint()
        {
            var transform = new Transform();
            transform.Parent = _midPoint.Parent;
            transform.Position = transform.Position + _midPoint.LocalPosition;
            return transform;
        }

        internal Transform CopyStartPoint()
        {
            var translationVector = GetLocalTranslationVector(100);
            var transform = new Transform();
            transform.Parent = _midPoint.Parent;
            transform.Position = transform.Position + _midPoint.LocalPosition;
            return transform;
        }

        internal GroupRenderItem GetInnerItemGroup()
        {
            return _innerGroup;
        }

        /// <summary>
        /// Transform origin in local coordinates
        /// Translated
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetInnerGroupTransformOriginTranslated()
        {
            return new Coordinate(_innerGroup.TransformOrigin.X + _innerGroup.TranslationOffset.Left, _innerGroup.TransformOrigin.Y + _innerGroup.TranslationOffset.Top);
        }
        /// <summary>
        /// Transform origin in local coordinates
        /// Translated
        /// </summary>
        /// <returns></returns>
        internal Transform GetInnerGroupWithTransformOriginTranslated()
        {
            Transform transform = new Transform();
            transform.Parent = _innerGroup.Bounds.Parent;
            transform.LocalPosition += new Vector2(_innerGroup.TransformOrigin.X + _innerGroup.TranslationOffset.Left, _innerGroup.TransformOrigin.Y + _innerGroup.TranslationOffset.Top);
            return transform;
        }

        //bool CanFitWithinEndPosition()
        //{
        //    var test = _startPoint;
        //    var mid = _midPoint;
        //    var end = _endPoint;
        //    var ctr = _circleCenter;
        //}

        internal BoundingBox GetBounds()
        {
            BoundingBox box = new BoundingBox(LargestWidthRectangle, LargestHeightRectangle);
            box.Parent = ExtremePoints.Parent;
            box.Left = ExtremePoints.Left;
            box.Top = ExtremePoints.Top;
            return box;
        }

        internal Transform GetCenterOfStartPointLine()
        {
            //var startToCenter = _startPoint.Position - _circleCenter.Position;
            //var startToCenterDirOnly = startToCenter / startToCenter.Length;

            //var half = (startToCenterDirOnly * -1) * 0.5d;
            //var halfPos = _startPoint.LocalPosition * half;

            //Transform transform = new Transform();
            //transform.Position = _startPoint.Position * half;
            //transform.Parent = _startPoint.Parent;

            return _startPoint;
        }

        

        internal Transform GetCenterOfEndPointLine()
        {
            //var ctrToEnd = _endPoint.LocalPosition - _circleCenter.LocalPosition;
            //var ctrToEndDirOnly = ctrToEnd / ctrToEnd.Length;

            //var half = ctrToEndDirOnly * 0.5d;
            //var halfPos = _circleCenter.LocalPosition * half;

            //Transform transform = new Transform();
            //transform.Parent = _innerGroup.Bounds;
            //transform.LocalPosition = halfPos;

            //return transform;
            return _endPoint;
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {

            throw new NotImplementedException();
        }

        //internal BoundingBox GetInnerGroupBounds()
        //{
        //    return _innerGroup.Bounds;
        //}
    }
}
