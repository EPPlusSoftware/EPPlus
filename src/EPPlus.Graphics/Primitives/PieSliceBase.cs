using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Graphics.Geometry;
using System.Text;

namespace EPPlus.Graphics.Primitives
{
    public class PieSliceBase
    {
        Vector2 _startPoint;
        Vector2 _endPoint;
        Vector2 _midPoint;

        public Vector2 StartPoint
        {
            get
            {
                //CalculateSliceArcPoints();
                return _startPoint;
            }
            private set
            {
                _startPoint = value;
            }
        }

        public Vector2 EndPoint
        {
            get
            {
                //CalculateSliceArcPoints();
                return _endPoint;
            }
            private set
            {
                _endPoint = value;
            }
        }

        public Vector2 MidPoint
        {
            get
            {
                //CalculateSliceArcPoints();
                return _midPoint;
            }
            private set
            {
                _midPoint = value;
            }
        }

        /// <summary>
        /// Vector from center of circle to outer middle of Pie-Slice
        /// This is the vector the slice is translated along in a pie-explosion
        /// </summary>
        public Vector2 ExplosionVector { get { return MidPoint - Circle.Center; } }

        public double StartDegrees { get; private set; }
        public double EndDegrees { get; private set; }

        internal double SweepAngleInDegrees { get { return EndDegrees - StartDegrees; } }

        public PrimitiveCircle Circle = null;

        internal PieSliceBase(double startDegrees, double endDegrees, PrimitiveCircle circle)
        {
            Circle = circle;
            Initialize(startDegrees, endDegrees);
        }

        internal PieSliceBase(double degrees, double prevSliceDegrees, double radius, Vector2 originPoint)
        {
            if (Circle == null)
            {
                Circle = new PrimitiveCircle(originPoint, radius);
            }

            Initialize(prevSliceDegrees, degrees + prevSliceDegrees);
        }


        void Initialize(double startDegrees, double endDegrees)
        {
            StartDegrees = startDegrees;
            EndDegrees = endDegrees;
            CalculateSliceArcPoints();
        }

        void CalculateSliceArcPoints()
        {
            _startPoint = Circle.GetPointOnCircle(StartDegrees);

            _endPoint = Circle.GetPointOnCircle(EndDegrees);

            //The degrees of the midpoint
            var halfDegrees = EndDegrees / 2;

            //We add prev at this point since we don't want to halve the previous angle only the current one
            _midPoint = Circle.GetPointOnCircle(halfDegrees + StartDegrees);

            CalculateWidthHeight(StartDegrees);
            CalculateLargestRectWithinCircleSegment();
        }

        bool ExistWithinRange(double target, double min, double max)
        {
            if (min < target && target < max)
            {
                return true;
            }
            return false;
        }

        public BoundingBox ExtremePoints;

        void CalculateWidthHeight(double prevSliceDegrees)
        {
            var endPointDegrees = prevSliceDegrees + SweepAngleInDegrees;
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
                maxY = Circle.Center.Y + Circle.Radius;
            }
            else
            {
                maxY = Math.Max(Circle.Center.Y, _endPoint.Y);
            }

            maxY = Math.Max(maxY, Circle.Center.Y);

            if (ExistWithinRange(180, startPointDegrees, endPointDegrees))
            {
                minX = Circle.Center.X - Circle.Radius;
            }
            else
            {
                minX = Math.Min(_startPoint.X, _endPoint.X);
            }

            minX = Math.Min(minX, Circle.Center.X);

            if (ExistWithinRange(270, startPointDegrees, endPointDegrees))
            {
                minY = Circle.Center.Y - Circle.Radius;
            }
            else
            {
                minY = Math.Min(_startPoint.Y, _endPoint.Y);
            }

            minY = Math.Min(minY, Circle.Center.Y);

            if (endPointDegrees < startPointDegrees || ExistWithinRange(0, startPointDegrees, endPointDegrees))
            {
                if (endPointDegrees > 270)
                {
                    maxX = Circle.Center.X;
                }
                else
                {
                    maxX = Circle.Center.X + Circle.Radius;
                }
            }
            else
            {
                maxX = Math.Max(_startPoint.X, _endPoint.X);
            }

            maxX = Math.Max(Circle.Center.X, maxX);


            ExtremePoints = new BoundingBox(minX, minY, maxX - minX, maxY - minY);
        }

        internal double LargestWidthRectangle { get; private set; }
        internal double LargestHeightRectangle { get; private set; }

        internal double ContentRectangleTop { get; private set; }
        internal double ContentRectangleLeft { get; private set; }

        //See Internal Docs: "Inscribed Rectangle" in Atlassian for details
        void CalculateLargestRectWithinCircleSegment()
        {
            //The degrees of the slice
            var myDegrees = SweepAngleInDegrees;

            ////Rotate back
            //myDegrees = myDegrees + 90;

            var boundedDegrees = myDegrees % 360d;

            if (SweepAngleInDegrees < 180d)
            {
                //Formula for largest (unrotated) rectangle within a circle-section
                //Calculate thetha = alpha/4
                var angleForTriangle = SweepAngleInDegrees / 4d;

                var angleForYTriangle = angleForTriangle + 0.64d;

                var yTriangle = (Math.Sin(GeoMathUtils.DegreesToRadians(angleForYTriangle)) * Circle.Radius);// add 1 for small rounding fault making too small
                var xTriangle = (Math.Cos(GeoMathUtils.DegreesToRadians(angleForTriangle)) * Circle.Radius);

                yTriangle *= 2d;

                //Normalize just in case
                var dirOnly = ExplosionVector / ExplosionVector.Length;

                LargestWidthRectangle = xTriangle;
                LargestHeightRectangle = yTriangle;


                var angleForTopLeftCalc = SweepAngleInDegrees / 2d;

                ContentRectangleLeft = yTriangle / Math.Tan(GeoMathUtils.DegreesToRadians(angleForTopLeftCalc));
                //Up is considered negative y direction and the center of the circle is local origin
                ContentRectangleTop = -yTriangle/2d;


                ////special case for quadrant -,-
                ////Y positive because of cartesian coord system
                //if (dirOnly.Y > 0 && dirOnly.X < 0)
                //{
                //    //As our starting point is not at the unit circle start. X and Y may need to be switched for width and height.
                //    //Our formula above assumes starting in unit circle and rotating upwards

                //    //We start at the "top" and rotate the other direction
                //    //Therefore:

                //    //since unit circle has been rotated width and height flip
                //    LargestWidthRectangle = yTriangle;
                //    LargestHeightRectangle = xTriangle;
                //}
                //else
                //{
                //    if (SweepAngleInDegrees < 90d)
                //    {
                //        LargestWidthRectangle = yTriangle;
                //        LargestHeightRectangle = xTriangle;
                //    }
                //    else
                //    {
                //        LargestWidthRectangle = xTriangle;
                //        LargestHeightRectangle = yTriangle;
                //    }
                //}
            }
            else
            {
                //Formula for largest (unrotated) rectangle within a semi-circle
                LargestWidthRectangle = Math.Sqrt(2d) * Circle.Radius;
                LargestHeightRectangle = (Math.Sqrt(2d) / 2d) * Circle.Radius;
            }
        }

    }
}
