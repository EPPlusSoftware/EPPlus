using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;

namespace EPPlus.Graphics
{
    internal class BoundingBox : Rect
    {
        internal Transform transform;

        private BoundingBox _parent = null;

        internal BoundingBox Parent { get { return _parent; } set { _parent = value; transform.Parent = value.transform; } }

        bool ClampedToParent = false;

        internal BoundingBox() : base()
        {
            transform = new Transform();
        }

        internal BoundingBox(double width, double height) : this()
        {
            Left = 0;
            Top = 0;
            Right = width;
            Bottom = height;
        }
        internal BoundingBox(double left, double top, double right, double bottom) : this()
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        /// <summary>
        /// Y pos (min)
        /// </summary>
        internal override double Top
        {
            get { return transform.LocalPosition.Y; }
            set
            {
                var tmpHeight = Height != 0 ? Height : 0;

                var currentPosition = transform.LocalPosition;
                currentPosition.Y = value;
                transform.LocalPosition = currentPosition;

                //Recalculate bottom position correctly
                if (tmpHeight != 0)
                {
                    Height = tmpHeight;
                }
            }
        }
        /// <summary>
        /// X pos (min)
        /// </summary>
        internal override double Left
        {
            get { return transform.LocalPosition.X; }
            set
            {
                var tmpWidth = Width != 0 ? Width : 0;

                var currentPosition = transform.LocalPosition;
                currentPosition.X = value;
                transform.LocalPosition = currentPosition;

                //Recalculate Right position correctly
                if (tmpWidth != 0)
                {
                    Width = tmpWidth;
                }
            }
        }

        /// <summary>
        /// If @ClampedToParent is true will not set value beyond parent
        /// </summary>
        internal override double Bottom { get => base.Bottom; set => SetBottom(value); }

        /// <summary>
        /// If @ClampedToParent is true will not set value beyond parent
        /// </summary>
        internal override double Right { get => base.Right; set => SetRight(value); }

        private void SetRight(double value)
        {
            if(ClampedToParent)
            {
                if(transform.Parent != null)
                {
                    var newValue = System.Math.Min(value, Parent.Right);
                    Right = newValue;
                }
            }
            Right = value;
        }

        private void SetBottom(double value)
        {
            if (ClampedToParent)
            {
                if (transform.Parent != null)
                {
                    var newValue = System.Math.Min(value, Parent.Bottom);
                    Bottom = newValue;
                }
            }
            Bottom = value;
        }

        //Quick-access to underlying transform

        /// <summary>
        /// Local position X
        /// X-position from parent transform position
        /// </summary>
        internal override double X { get { return Left; } set { Left = value; } }

        /// <summary>
        /// Local position Y
        /// Y-position from parent transform position
        /// </summary>
        internal override double Y { get { return Top; } set { Top = value; } }

        //Gets global position x and y
        internal double GlobalX { get { return transform.Position.X; } }
        internal double GlobalY { get { return transform.Position.Y; } }
    }
}
