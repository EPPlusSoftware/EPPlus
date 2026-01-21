using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    internal class BoundingBox : Rect
    {
        internal Transform Transform { get; }

        private BoundingBox _parent = null;

        internal BoundingBox Parent { get { return _parent; } set { _parent = value; Transform.Parent = value.Transform; } }

        bool ClampedToParent = false;

        internal BoundingBox() : base()
        {
            Transform = new Transform();
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
            get { return Transform.LocalPosition.Y; }
            set
            {
                var tmpHeight = Height != 0 ? Height : 0;

                var currentPosition = Transform.LocalPosition;
                currentPosition.Y = value;
                Transform.LocalPosition = currentPosition;

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
            get { return Transform.LocalPosition.X; }
            set
            {
                var tmpWidth = Width != 0 ? Width : 0;

                var currentPosition = Transform.LocalPosition;
                currentPosition.X = value;
                Transform.LocalPosition = currentPosition;

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

        internal override double Width {
            get
            {
                return Right - Left;
            }
            set
            {
                Right = Left + value;
            }
        
        }

        internal override double Height
        {
            get
            {
                return Bottom - Top;
            }
            set
            {
                Bottom = Top + value;
            }
        }

        private void SetRight(double value)
        {
            if(ClampedToParent)
            {
                if(Transform.Parent != null)
                {
                    var newValue = System.Math.Min(value, Parent.Right);
                    base.Right = newValue;
                }
            }
            base.Right = value;
        }

        private void SetBottom(double value)
        {
            if (ClampedToParent)
            {
                if (Transform.Parent != null)
                {
                    var newValue = System.Math.Min(value, Parent.Bottom);
                    base.Bottom = newValue;
                }
            }
            base.Bottom = value;
        }

        //Quick-access to underlying Transform

        /// <summary>
        /// Local position X
        /// X-position from parent transform position
        /// </summary>
        internal override double X { get { return Left; } set { Left = value; } }

        /// <summary>
        /// Local position Y
        /// Y-position from parent Transform position
        /// </summary>
        internal override double Y { get { return Top; } set { Top = value; } }

        //Gets global position x and y
        internal double GlobalX { get { return Transform.Position.X; } }
        internal double GlobalY { get { return Transform.Position.Y; } }
    }
}
