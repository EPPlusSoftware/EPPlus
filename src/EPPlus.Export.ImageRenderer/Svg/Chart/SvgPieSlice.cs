using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using System;

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

        public SvgPieSlice(DrawingBase renderer, BoundingBox parent, Point circleCenter, double radius, double percentOfPie, double prevSliceDegrees) : base(renderer, parent)
        {
            _radius = radius;
            _percent = percentOfPie;
            //How many degrees that percentage is out of 360
            Degrees = _percent * 360d;

            _innerGroup = new SvgGroupItemNew(renderer, parent, 0, circleCenter);
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
        }

        internal void ImportPathData(BoundingBox plotAreaBounds, BoundingBox globalAreaBounds, double sliceScaleFactor, double explosionOfPoint, double pieExplosion, int position)
        {
            _slicePath = new SvgRenderPathItem(DrawingRenderer, plotAreaBounds);

            #region calculate global variables
            //Path commands are based on percent of the Pixel height and Width
            //We must supply coords in a correct percentage of the Whole Chart
            var w = plotAreaBounds.Width;
            var h = plotAreaBounds.Height;

            var cx = (w / 2);
            var cy = (h / 2);

            var radiusXAspectRatioPercent = _radius / w;
            var radiusYAspectRatioPercent = _radius / h;

            var circleCenterWorld = _circleCenter.Position;

            var cxPercentOfTotal = circleCenterWorld.X / globalAreaBounds.Width;
            var cyPercentOfTotal = circleCenterWorld.Y / globalAreaBounds.Height;

            var moveCenter = new PathCommands(PathCommandType.Move, _slicePath, cxPercentOfTotal, cyPercentOfTotal);

            var lastPosX = _startPoint.Position.X / globalAreaBounds.Width;
            var lastPosY = _startPoint.Position.Y / globalAreaBounds.Height;
            Coordinate startPointGlobalPercentage = new Coordinate(lastPosX, lastPosY);

            var lineToStart = new PathCommands(PathCommandType.Line, _slicePath, startPointGlobalPercentage.X, startPointGlobalPercentage.Y);

            var worldEnd = _endPoint.TransformPointToWorld(_endPoint.LocalPosition);

            var arcCommand = new PathCommands(PathCommandType.Arc, _slicePath, new double[] { radiusXAspectRatioPercent, radiusYAspectRatioPercent, 0, Degrees > 180 ? 1 : 0, 1, worldEnd.X / w, worldEnd.Y / h });
            var end = new PathCommands(PathCommandType.End, _slicePath, worldEnd.X / w, worldEnd.Y / h);


            //Get maximum local extreme values from global values
            var localMax = GetTranslationMaxLocal(globalAreaBounds.Width, globalAreaBounds.Height);
            var localMin = GetTranslationMinLocal(globalAreaBounds.Width, globalAreaBounds.Height);
            #endregion

            //Scale the inner group to pie explosion
            _innerGroup.Scale = new Coordinate(sliceScaleFactor, sliceScaleFactor);
            CalculatePointExplosion(explosionOfPoint, pieExplosion, localMax, localMin);

            _slicePath.Commands.Add(moveCenter);
            _slicePath.Commands.Add(lineToStart);
            //_slicePath.Commands.Add(arcCommand);
            //_slicePath.Commands.Add(end);
        }

        internal void ImportStlyeInfo(ExcelChartDataPoint dp, ExcelPieChart chartType)
        {
            _slicePath.SetDrawingPropertiesFill(dp.Fill, chartType.StyleManager.Style.DataPoint.FillReference.Color);
            _slicePath.SetDrawingPropertiesBorder(dp.Border, chartType.StyleManager.Style.DataPoint.BorderReference.Color, true);
            _slicePath.SetDrawingPropertiesEffects(dp.Effect);
            _slicePath.BorderWidth = dp.Border.Width;
        }

        internal void AppendGroupItem(SvgGroupItemNew group)
        {
            _innerGroup.AddChildItem(_slicePath);
            group.AddChildItem(_innerGroup);
        }

        Graphics.Math.Vector2 GetTranslationMaxLocal(double globalWidth, double globalHeight)
        {
            var worldPositionTransformOrigin = _midPoint.Position;

            //Calculate extremes 
            Point worldMax = new Point(
                globalWidth - worldPositionTransformOrigin.X,
                globalHeight - worldPositionTransformOrigin.Y);

            return _innerGroup.Position.TransformPointToLocal(worldMax.Position);
        }

        Graphics.Math.Vector2 GetTranslationMinLocal(double globalWidth, double globalHeight)
        {
            var worldPositionTransformOrigin = _midPoint.Position;
            //Calculate extremes 
            Point worldMin = new Point(-worldPositionTransformOrigin.X, -worldPositionTransformOrigin.Y);

            var localMin = _innerGroup.Position.TransformPointToLocal(worldMin.Position);
            return localMin;
        }

        Graphics.Math.Vector2 GetLocalTranslationVector(double explosionOfPoint, double pieExplosion)
        {
            //Get point explosion value
            var pointExplosion = explosionOfPoint == int.MinValue ? 0 : explosionOfPoint;

            var transformOriginLocal = new Graphics.Math.Vector2(_innerGroup.TransformOrigin.X, _innerGroup.TransformOrigin.Y);

            //Get directional vector (in local coords but does not matter since we make it directional)
            Graphics.Math.Vector2 pieDirection = transformOriginLocal - _circleCenter.Position;

            //normalize the pieDirection vector so that it is percentual and with lenght == 1
            pieDirection = pieDirection / pieDirection.Length;

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

            //Get distance/length to move along vector
            Graphics.Math.Vector2 LocalTranslationVector = ptFactoredDirection * _radius;
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
            var finalTranslation = GetFinalLocalTranslation(LocalTranslationVector, localMax,localMin);
            
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
            point.Parent = _innerGroup.Bounds;
            point.Left = xPoint;
            point.Top = yPoint;

            return point;
        }

        public override RenderItemType Type => RenderItemType.Group;

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }
    }
}
