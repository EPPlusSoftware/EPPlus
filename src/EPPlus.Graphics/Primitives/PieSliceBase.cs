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
                CalculateSliceArcPoints();
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
                CalculateSliceArcPoints();
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
                CalculateSliceArcPoints();
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

        internal double StartDegrees;
        internal double EndDegrees;

        internal double SweepAngleInDegrees { get { return EndDegrees - StartDegrees; } }

        internal PrimitiveCircle Circle = null;

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

            Initialize(degrees, prevSliceDegrees);
        }


        void Initialize(double startDegrees, double endDegrees)
        {
            StartDegrees = startDegrees;
            EndDegrees = endDegrees;
            CalculateSliceArcPoints();
        }

        void CalculateSliceArcPoints()
        {
            StartPoint = Circle.GetPointOnCircle(StartDegrees);

            EndPoint = Circle.GetPointOnCircle(EndDegrees + StartDegrees);

            //The degrees of the midpoint
            var halfDegrees = EndDegrees / 2;

            //We add prev at this point since we don't want to halve the previous angle only the current one
            MidPoint = Circle.GetPointOnCircle(halfDegrees + StartDegrees);
        }
    }
}
