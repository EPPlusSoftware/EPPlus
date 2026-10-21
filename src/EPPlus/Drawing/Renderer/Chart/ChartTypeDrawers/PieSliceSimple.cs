using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlus.Graphics.Primitives;
using EPPlus.Graphics.TransformPrimitives;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Drawing.Renderer.Chart.ChartTypeDrawers
{
    internal class PieSliceSimple : PieSliceTransform
    {
        internal PieSliceSimple(PieSliceBase sliceBase, TransformCircle parent = null) : base(sliceBase, parent)
        {
            CalculateWidthHeight(sliceBase.StartDegrees);
        }

        PathRenderItem _slicePath;

        internal void ImportPathData(BoundingBox pathParent, int position)
        {
            GroupRenderItem group = new GroupRenderItem(pathParent);

            group.Style.FillColor = "Purple";

            _slicePath = new PathRenderItem(group);
            _slicePath.Style.BorderWidth = 5d;

            //Calculate path commands
            var moveCenter = new PathCommand(PathCommandType.Move, Circle.Left, Circle.Top);
            var lineToStart = new PathCommand(PathCommandType.Line, StartPoint.Left, StartPoint.Top);

            var arcCommand = new PathCommand(PathCommandType.Arc, new double[] { Circle.Radius, Circle.Radius, 0, Degrees > 180 ? 1 : 0, 1, EndPoint.Left, EndPoint.Top });
            var end = new PathCommand(PathCommandType.End, EndPoint.Left, EndPoint.Top);

            //Add the actual commands
            _slicePath.Commands.Add(moveCenter);
            _slicePath.Commands.Add(lineToStart);
            _slicePath.Commands.Add(arcCommand);

            GroupRenderItem debugGroup = new GroupRenderItem(pathParent);
            debugGroup.Style.FillColor = "darkRed";

            AddDebugLines(moveCenter, debugGroup);

            _slicePath.Commands.Add(end);

        }

        BoundingBox ExtremePoints = new BoundingBox();

        bool ExistWithinRange(double target, double min, double max)
        {
            if (min < target && target < max)
            {
                return true;
            }
            return false;
        }

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

            if (/*endPointDegrees < startPointDegrees || */ExistWithinRange(90, startPointDegrees, endPointDegrees))
            {
                maxY = Circle.Top + Circle.Radius;
            }
            else
            {
                maxY = Math.Max(StartPoint.Top, EndPoint.Top);
            }

            maxY = Math.Max(maxY, Circle.Top);

            if (ExistWithinRange(180, startPointDegrees, endPointDegrees))
            {
                minX = Circle.Left - Circle.Radius;
            }
            else
            {
                minX = Math.Min(StartPoint.Left, EndPoint.Left);
            }

            minX = Math.Min(minX, Circle.Left);

            if (ExistWithinRange(270, startPointDegrees, endPointDegrees))
            {
                minY = Circle.Top - Circle.Radius;
            }
            else
            {
                minY = Math.Min(StartPoint.Top, EndPoint.Top);
            }

            minY = Math.Min(minY, Circle.Top);

            if (endPointDegrees < startPointDegrees || ExistWithinRange(0, startPointDegrees, endPointDegrees))
            {
                if (endPointDegrees > 270)
                {
                    maxX = Circle.Left;
                }
                else
                {
                    maxX = Circle.Left + Circle.Radius;
                }
            }
            else
            {
                maxX = Math.Max(StartPoint.Left, EndPoint.Left);
            }

            maxX = Math.Max(Circle.Left, maxX);


            ExtremePoints = new BoundingBox(minX, minY, maxX - minX, maxY - minY);
            //ExtremePoints.Position = Circle.Position;
            //ExtremePoints.Parent = _innerGroup;
        }

        List<RenderItem> DebugItems;
        PathRenderItem _debugBoundsPath;


        /// <summary>
        /// Get copy of start point of slice in local coordinates
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetStartPointPositionLocal()
        {
            return new Coordinate(StartPoint.Left, StartPoint.Top);
        }
        /// <summary>
        /// Get copy of mid point of slice in local coordinates
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetMidPointLocal()
        {
            return new Coordinate(MidPoint.Left, MidPoint.Top);
        }
        /// <summary>
        /// Get copy of end point of slice in local coordinates
        /// </summary>
        /// <returns></returns>
        internal Coordinate GetEndPointLocal()
        {
            return new Coordinate(EndPoint.LocalPosition.X, EndPoint.LocalPosition.Y);
        }

        RectRenderItem _debugCircleCenter;

        /// <summary>
        /// Adds line from center to outer mid point (transform-origin) and to end point
        /// AKA line along scale/explosion vector
        /// </summary>
        /// <param name="moveCenter"></param>
        private void AddDebugLines(PathCommand moveCenter, BoundingBox bounds)
        {
            DebugItems = new List<RenderItem>();


            var inscribedRect = new RectRenderItem(bounds.Parent);
            inscribedRect.Width = LargestInscribedRectangle.Width;
            inscribedRect.Height = LargestInscribedRectangle.Height;

            inscribedRect.Style.FillColor = "blue";
            inscribedRect.Style.FillOpacity = 0.2d;

            _debugBoundsPath = new PathRenderItem(bounds);

            //Render bounds for slice
            _debugBoundsPath = new PathRenderItem(bounds);
            _debugBoundsPath.Style.BorderColor = "red";
            _debugBoundsPath.Style.FillColor = "transparent";
            _debugBoundsPath.Style.BorderWidth = 3;
            var moveCenterDebug = new PathCommand(PathCommandType.Move, ExtremePoints.Left, ExtremePoints.Top);
            _debugBoundsPath.Commands.Add(moveCenterDebug);

            //Draw extremes/bounds
            //var lineToTopLeft = new PathCommands(PathCommandType.Line, ExtremePoints.Left, ExtremePoints.Top);
            var lineToTopRight = new PathCommand(PathCommandType.Line, ExtremePoints.Right, ExtremePoints.Top);

            var lineToBottomRight = new PathCommand(PathCommandType.Line, ExtremePoints.Right, ExtremePoints.Bottom);
            var lineToBottomLeft = new PathCommand(PathCommandType.Line, ExtremePoints.Left, ExtremePoints.Bottom);
            var end = new PathCommand(PathCommandType.End, ExtremePoints.Left, ExtremePoints.Top);

            _debugBoundsPath.Commands.Add(lineToTopRight);
            _debugBoundsPath.Commands.Add(lineToBottomRight);
            _debugBoundsPath.Commands.Add(lineToBottomLeft);
            _debugBoundsPath.Commands.Add(end);

            DebugItems.Add(_debugBoundsPath);

            //Render green dot at circle center
            _debugCircleCenter = GenerateDebugPoint(bounds, new Coordinate(Circle.Left, Circle.Top), "green");

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

            var rect = new RectRenderItem(parent) { Left = l + point.X, Top = t + point.Y, Width = w, Height = h };
            rect.Style.FillColor = fillColor;
            return rect;
        }
    }
}
