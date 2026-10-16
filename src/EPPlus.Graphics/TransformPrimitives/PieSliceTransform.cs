using EPPlus.Graphics.Geometry;
using EPPlus.Graphics.Primitives;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics.TransformPrimitives
{
    public class PieSliceTransform : BoundingBox
    {
        //PieSliceBase sliceBase;

        public TransformCircle Circle { get; private set; }

        TranformPoint _startPoint;
        TranformPoint _endPoint;
        TranformPoint _midPoint;

        public TranformPoint StartPoint
        {
            get
            {
                return _startPoint;
            }
            private set
            {
                _startPoint = value;
            }
        }

        public TranformPoint EndPoint
        {
            get
            {
                return _endPoint;
            }
            private set
            {
                _endPoint = value;
            }
        }

        public TranformPoint MidPoint
        {
            get
            {
                return _midPoint;
            }
            private set
            {
                _midPoint = value;
            }
        }

        public Vector2 ExplosionVector { get { return MidPoint.LocalPosition - Circle.LocalPosition; } }

        public Vector2 ExplosionDir { get { return ExplosionVector / ExplosionVector.Length; } }

        public double StartDegrees;

        public double EndDegrees;

        public double Degrees;

        public BoundingBox LargestInscribedRectangle { get; private set; }

        internal protected PieSliceTransform(PieSliceBase sliceBase, TransformCircle parent = null) : base()
        {
            Left = sliceBase.ExtremePoints.Left;
            Top = sliceBase.ExtremePoints.Top;
            Width = sliceBase.ExtremePoints.Right - Left;
            Height = sliceBase.ExtremePoints.Bottom - Top;

            if(parent == null)
            {
                Circle = new TransformCircle(sliceBase.Circle);
                Parent = Circle;
            }
            else
            {
                Circle = parent;
                Parent = Circle;
            }

            StartPoint = new TranformPoint(sliceBase.StartPoint.X, sliceBase.StartPoint.Y);
            MidPoint = new TranformPoint(sliceBase.MidPoint.X, sliceBase.MidPoint.Y);
            EndPoint = new TranformPoint(sliceBase.EndPoint.X, sliceBase.EndPoint.Y);

            StartPoint.Parent = Circle;
            MidPoint.Parent = Circle;
            EndPoint.Parent = Circle;

            StartDegrees = sliceBase.StartDegrees;
            EndDegrees = sliceBase.EndDegrees;
            Degrees = sliceBase.SweepAngleInDegrees;

            ////Top and bottom positions could potentially be added
            ////But as of right now this is only used to determine if it is possible to place a text of a certain size inside the slice or not.
            //LargestInscribedRectangle = new BoundingBox(sliceBase.LargestWidthRectangle, sliceBase.LargestHeightRectangle);
            LargestInscribedRectangle = new BoundingBox(sliceBase.ContentRect.Left, sliceBase.ContentRect.Top, sliceBase.ContentRect.Right - sliceBase.ContentRect.Left, sliceBase.ContentRect.Bottom - sliceBase.ContentRect.Top);
        }
    }
}
