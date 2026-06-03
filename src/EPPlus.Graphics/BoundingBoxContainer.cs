using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    internal class BoundingBoxContainer : BoundingBox
    {
        internal double MarginLeft
        {
            get
            {
                return InnerBounds.Left;
            }
            set
            {
                InnerBounds.Left = value;
            }
        }

        internal double MarginTop
        {
            get
            {
                return InnerBounds.Top;
            }
            set
            {
                InnerBounds.Top = value;
            }
        }

        internal double MarginRight
        {
            get
            {
                return GetInnerRight();
            }
            set
            {
                InnerBounds.Width = Width - value;
            }
        }
        internal double MarginBottom
        {
            get
            {
                return GetInnerBottom();
            }
            set
            {
                InnerBounds.Height = Height - value;
            }
        }

        BoundingBox InnerBounds;

        internal BoundingBoxContainer(BoundingBox innerBounds) : base() 
        {
            InnerBounds = innerBounds;
        }

        //internal BoundingBoxContainer(double l, double t, double r, double b) : base(l, t, r, b)
        //{
        //}

        internal double GetInnerLeft()
        {
            return Left + MarginLeft;
        }

        internal double GetInnerTop()
        {
            return Top + MarginTop;
        }

        internal double GetInnerBottom()
        {
            return Bottom - MarginBottom;
        }

        internal double GetInnerRight()
        {
            return Right - MarginRight;
        }

        //public BoundingBox GetInnerBounds()
        //{
        //    var textArea = new BoundingBox();
        //    textArea.Left = GetInnerLeft();
        //    textArea.Top = GetInnerTop();

        //    textArea.Width = GetInnerWidth();
        //    textArea.Height = GetInnerHeight();

        //    return textArea;
        //}

        public double GetInnerWidth()
        {
            return GetInnerRight() - GetInnerLeft();
        }

        public double GetInnerHeight()
        {
            return GetInnerBottom() - GetInnerTop();
        }
    }
}
