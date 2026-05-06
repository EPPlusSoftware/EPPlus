using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.IO.Pipes;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class SvgPieSlice : SvgRenderItem
    {
        double _radius;
        SvgGroupItemNew _innerGroup;
        Point _circleCenter;

        /// <summary>
        /// How many percent of the pie this represents
        /// </summary>
        double _percent;

        internal double Degrees { get; private set; }

        Point _startPoint;
        Point _midPoint;
        Point _endPoint;

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
            return new Coordinate(_endPoint.Left, _endPoint.Top);
        }

        SvgRenderPathItem _slicePath;
        SvgRenderPathItem _debugBoundsPath;

        bool ExistWithinRange(double target, double min, double max)
        {
            if(min < target && target < max)
            {
                return true;
            }
            return false;
        }

        internal BoundingBox ExtremePoints { get; set; }

        void CalculateWidthHeight(double prevSliceDegrees)
        {
            var endPointDegrees = prevSliceDegrees + Degrees;
            if(endPointDegrees < 0)
            {
                endPointDegrees = 360 + endPointDegrees;
            }

            var startPointDegrees = prevSliceDegrees;
            if(startPointDegrees < 0)
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

            if(ExistWithinRange(270, startPointDegrees, endPointDegrees))
            {
                minY = _circleCenter.Top - _radius;
            }
            else
            {
                minY = Math.Min(_startPoint.Top, _endPoint.Top);
            }

            minY = Math.Min(minY, _circleCenter.Top);

            if (ExistWithinRange(360, startPointDegrees, endPointDegrees) || ExistWithinRange(0, startPointDegrees, endPointDegrees))
            {
                maxX = _circleCenter.Left + _radius;
            }
            else
            {
                maxX = Math.Max(_startPoint.Left, _endPoint.Left);
            }

            maxX = Math.Max(_circleCenter.Left, maxX);


            ExtremePoints = new BoundingBox(minX, minY, maxX - minX, maxY - minY);
            ExtremePoints.Parent = Bounds;
            //if(Degrees > 0)
            //{

            //}
            //if(Degrees)
            //double xMax = Math.Max(_startPoint.Left, _endPoint.Left);
            //xMax = Math.Max(xMax, _midPoint.Left);

            //double yMax = Math.Max(_startPoint.Top, _endPoint.Top);
            //yMax = Math.Max(yMax, _midPoint.Top);
            //double yMax;
            //double xMin;
            //double yMin;
        }

        private double _sliceScaleFactor = 1d;
        private double _scaledRadius { get{ return _radius * _sliceScaleFactor; } }

        private void CalculateExplosionDir()
        {
            var transformOriginLocal = new Graphics.Math.Vector2(_innerGroup.TransformOrigin.X, _innerGroup.TransformOrigin.Y);

            //Get directional vector (in local coords but does not matter since we make it directional)
            Graphics.Math.Vector2 pieDirection = transformOriginLocal - _circleCenter.LocalPosition;

            //normalize the pieDirection vector so that it is percentual and with lenght == 1
            pieDirection = pieDirection / pieDirection.Length;

            CtrToOuterMidDir = new Graphics.Math.Vector2(pieDirection.X, pieDirection.Y);
        }

        //internal Graphics.Math.Vector2 GetVectorCtrToEnd()
        //{
        //    //_circleCenter + _sc
        //}

        public SvgPieSlice(DrawingBase renderer, BoundingBox parent, Point circleCenter, double radius, double percentOfPie, double prevSliceDegrees) : base(renderer, parent)
        {
            _radius = radius;
            _percent = percentOfPie;
            //How many degrees that percentage is out of 360
            Degrees = _percent * 360d;

            _innerGroup = new SvgGroupItemNew(renderer, parent, 0, circleCenter);
            _innerGroup.Bounds.Parent = _innerGroup.Position;
            _circleCenter = circleCenter;

            _startPoint = CalculateLocalPointOnCircle(prevSliceDegrees);

            //The degrees of the midpoint
            var halfDegrees = Degrees / 2;

            _endPoint = CalculateLocalPointOnCircle(Degrees + prevSliceDegrees);

            //We add prev at this point since we don't want to halve the previous angle only the current one
            _midPoint = CalculateLocalPointOnCircle(halfDegrees + prevSliceDegrees);

            //We must calculate transforms from the outer midpoint.
            //This is to ensure that point never leaves the parent container
            _innerGroup.TransformOrigin = GetMidPointLocal();

            CalculateExplosionDir();

            CalculateWidthHeight(prevSliceDegrees);
        }

        internal void ImportPathData(BoundingBox plotAreaBounds, BoundingBox globalAreaBounds, double sliceScaleFactor, double explosionOfPoint, double pieExplosion, int position)
        {
            _slicePath = new SvgRenderPathItem(DrawingRenderer, plotAreaBounds);

            _slicePath.BorderWidth = 5;

            //Calculate path commands
            var moveCenter = new PathCommands(PathCommandType.Move, _slicePath, _circleCenter.Left, _circleCenter.Top);
            var lineToStart = new PathCommands(PathCommandType.Line, _slicePath, _startPoint.Left, _startPoint.Top);

            var arcCommand = new PathCommands(PathCommandType.Arc, _slicePath, new double[] { _radius, _radius, 0, Degrees > 180 ? 1 : 0, 1, _endPoint.Left, _endPoint.Top});
            var end = new PathCommands(PathCommandType.End, _slicePath, _endPoint.Left, _endPoint.Top);

            //Get max and min values
            var localMax = GetTranslationMaxLocal(globalAreaBounds.Width, globalAreaBounds.Height);
            var localMin = GetTranslationMinLocal(globalAreaBounds.Width, globalAreaBounds.Height);

            _sliceScaleFactor = sliceScaleFactor;
            //Translate and scale path
            _innerGroup.Scale = new Coordinate(_sliceScaleFactor, _sliceScaleFactor);
            CalculatePointExplosion(explosionOfPoint, pieExplosion, localMax, localMin);

            //Add the actual commands
            _slicePath.Commands.Add(moveCenter);
            _slicePath.Commands.Add(lineToStart);
            _slicePath.Commands.Add(arcCommand);

            //Visualize all points
            //AddDebugLines(moveCenter, plotAreaBounds);

            _slicePath.Commands.Add(end);

        }

        /// <summary>
        /// Adds line from center to outer mid point (transform-origin) and to end point
        /// AKA line along scale/explosion vector
        /// </summary>
        /// <param name="moveCenter"></param>
        private void AddDebugLines(PathCommands moveCenter, BoundingBox bounds)
        {
            //var lineToMidPoint = new PathCommands(PathCommandType.Line, _slicePath, _midPoint.Left, _midPoint.Top);
            //var lineToEnd = new PathCommands(PathCommandType.Line, _slicePath, _endPoint.Left, _endPoint.Top);
            //_slicePath.Commands.Add(moveCenter);
            //_slicePath.Commands.Add(lineToMidPoint);
            //_slicePath.Commands.Add(moveCenter);
            //_slicePath.Commands.Add(lineToEnd);
            _debugBoundsPath = new SvgRenderPathItem(DrawingRenderer, bounds);
            _debugBoundsPath.BorderColor = "red";
            _debugBoundsPath.FillColor = "transparent";
            _debugBoundsPath.BorderWidth = 3;
            var moveCenterDebug = new PathCommands(PathCommandType.Move, _debugBoundsPath, _circleCenter.Left, _circleCenter.Top);
            _debugBoundsPath.Commands.Add(moveCenterDebug);

            //Draw extremes/bounds
            var lineToTopLeft = new PathCommands(PathCommandType.Line, _debugBoundsPath, ExtremePoints.Left, ExtremePoints.Top);
            var lineToTopRight = new PathCommands(PathCommandType.Line, _debugBoundsPath, ExtremePoints.Right, ExtremePoints.Top);

            //var sliceCenter = GetSliceShapeCenterLocal();
            //var lineToSliceCenter = new PathCommands(PathCommandType.Line, _debugBoundsPath, sliceCenter.Left, sliceCenter.Top);

            var lineToBottomRight = new PathCommands(PathCommandType.Line, _debugBoundsPath, ExtremePoints.Right, ExtremePoints.Bottom);
            var lineToBottomLeft = new PathCommands(PathCommandType.Line, _debugBoundsPath, ExtremePoints.Left, ExtremePoints.Bottom);
            var end = new PathCommands(PathCommandType.End, _debugBoundsPath, ExtremePoints.Left, ExtremePoints.Top);
            _debugBoundsPath.Commands.Add(lineToTopLeft);
            _debugBoundsPath.Commands.Add(lineToTopRight);
            //_debugBoundsPath.Commands.Add(lineToSliceCenter);
            _debugBoundsPath.Commands.Add(lineToBottomRight);
            _debugBoundsPath.Commands.Add(lineToBottomLeft);
            //_debugBoundsPath.Commands.Add(end);

        }

        internal void ImportStlyeInfo(ExcelChartDataPoint dp, ExcelPieChart chartType)
        {
            _slicePath.SetDrawingPropertiesFill(dp.Fill, chartType.StyleManager.Style.DataPoint.FillReference.Color);
            _slicePath.SetDrawingPropertiesBorder(dp.Border, chartType.StyleManager.Style.DataPoint.BorderReference.Color, true);
            _slicePath.SetDrawingPropertiesEffects(dp.Effect);
        }

        internal void AppendGroupItem(SvgGroupItemNew group)
        {
            _innerGroup.AddChildItem(_slicePath);
            //_innerGroup.AddChildItem(_debugBoundsPath);
            group.AddChildItem(_innerGroup);
        }

        Graphics.Math.Vector2 GetTranslationMaxLocal(double globalWidth, double globalHeight)
        {
            var worldPositionTransformOrigin = _midPoint.Position;

            //Calculate extremes 
            Point localMax = new Point(
                globalWidth - worldPositionTransformOrigin.X,
                globalHeight - worldPositionTransformOrigin.Y);

            return localMax.LocalPosition;
        }

        Graphics.Math.Vector2 GetTranslationMinLocal(double globalWidth, double globalHeight)
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
        internal Graphics.Math.Vector2 CtrToOuterMidDir { get; private set; }

        /// <summary>
        /// Gets the whole vector with length to the end
        /// </summary>
        /// <returns></returns>
        internal Graphics.Math.Vector2 GetWholeVectorCenterToMid()
        {
            var translationVector = GetLocalTranslationVector(50);
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
        Graphics.Math.Vector2 GetLocalTranslationVector(double percentTowardsEndPoint)
        {
            if(percentTowardsEndPoint < 0 || percentTowardsEndPoint > 100)
            {
                throw new InvalidOperationException($"input: '{percentTowardsEndPoint}' invalid. Must be between 0 and 100");
            }

            var moveFactor = (double)percentTowardsEndPoint / 100d;
            var moveFactoredDirection = moveFactor * CtrToOuterMidDir;

            //Get distance/length to move along vector. We translate according to the scaled down radius
            Graphics.Math.Vector2 LocalTranslationVector = moveFactoredDirection * _scaledRadius;
            return LocalTranslationVector;
        }

        Graphics.Math.Vector2 GetLocalTranslationVector(double explosionOfPoint, double pieExplosion)
        {
            //Get point explosion value
            var pointExplosion = explosionOfPoint == int.MinValue ? 0 : explosionOfPoint;

            var pieDirection = new Graphics.Math.Vector2(CtrToOuterMidDir.X, CtrToOuterMidDir.Y);

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
            Graphics.Math.Vector2 LocalTranslationVector = ptFactoredDirection * _scaledRadius;
            return LocalTranslationVector;
        }

        Coordinate GetFinalLocalTranslation(Graphics.Math.Vector2 LocalTranslationVector, Graphics.Math.Vector2 localMax, Graphics.Math.Vector2 localMin)
        {
            Coordinate lengthPoint = new Coordinate(0,0);

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

                //translationLeft = Math.Min(translateX, maxTranslationXWorld);
                //translationTop = Math.Min(translateY, maxTranslationY);

                //translationLeft = Math.Max(translationLeft, minTranslationXWorld);
                //translationTop = Math.Max(translationTop, minTranslationYWorld);
            }

            return new Coordinate(translationLeft, translationTop);
        }

        void CalculatePointExplosion(double explosionOfPoint, double pieExplosion, Graphics.Math.Vector2 localMax, Graphics.Math.Vector2 localMin)
        {
            //Get distance/length to move along vector
            Graphics.Math.Vector2 LocalTranslationVector = GetLocalTranslationVector(explosionOfPoint, pieExplosion);
            var finalTranslation = GetFinalLocalTranslation(LocalTranslationVector, localMax, localMin);
            
            _innerGroup.Position.Left = finalTranslation.X;
            _innerGroup.Position.Top = finalTranslation.Y;
        }

        Point CalculateLocalPointOnCircle(double degrees)
        {
            var angleRadians = MConverter.DegreesToRadians(degrees);

            var xPoint = _circleCenter.Left + ( _radius * Math.Cos(angleRadians));
            var yPoint = _circleCenter.Top + (_radius * Math.Sin(angleRadians));

            var point = new Point();

            //Ensure the cx/cy offset
            point.Parent = _innerGroup.Position.Parent;
            point.Left = xPoint;
            point.Top = yPoint;

            return point;
        }

        internal Coordinate GetInnerGroupTransformOrigin()
        {
            return new Coordinate(_innerGroup.TransformOrigin.X, _innerGroup.TransformOrigin.Y);
        }

        public override RenderItemType Type => RenderItemType.Group;

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }
    }
}
