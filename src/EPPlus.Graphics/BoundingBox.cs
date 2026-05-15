using EPPlus.Graphics.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    public class BoundingBox : Transform
    {
        public BoundingBox() : base()
        {
        }

        public BoundingBox(double width, double height) : base(0, 0, width, height)
        {
        }
        public BoundingBox(double left, double top, double width, double height) : base(left, top, width, height)
        {
        }

        /// <summary>
        /// Y pos (min)
        /// </summary>
        public double Top
        {
            get { return LocalPosition.Y; }
            set
            {
                LocalPosition = new Vector2(LocalPosition.X, value);
                //var tmpHeight = Height != 0 ? Height : 0;

                //var currentPosition = Transform.LocalPosition;
                //currentPosition.Y = value;
                //Transform.LocalPosition = currentPosition;

                ////Recalculate bottom position correctly
                ////if (tmpHeight != 0)
                ////{
                //    //Height = tmpHeight;
                //    Bottom = Top + tmpHeight;
                ////}
            }
        }
        /// <summary>
        /// X pos (min)
        /// </summary>
        public double Left
        {
            get { return LocalPosition.X; }
            set
            {
                //var tmpWidth = Width != 0 ? Width : 0;

                //var currentPosition = Transform.LocalPosition;
                //currentPosition.X = value;
                //Transform.LocalPosition = currentPosition;

                //Recalculate Right position correctly
                //if (tmpWidth != 0)
                //{
                //Right = Left + tmpWidth;
                //Width = tmpWidth;
                //}
                LocalPosition = new Vector2(value, LocalPosition.Y);
            }
        }

        /// <summary>
        /// If @ClampedToParent is true will not set value beyond parent
        /// </summary>
        public double Bottom
        {
            get
            {
                return LocalPosition.Y + Size.Y;
            }
        }

        /// <summary>
        /// If @ClampedToParent is true will not set value beyond parent
        /// </summary>
        public double Right
        {
            get
            {
                return LocalPosition.X + Size.X;
            }
        }
        public virtual double Width
        {
            get
            {
                return Size.X;
            }
            set
            {
                Size = new Vector2(value, Size.Y);
            }
        }

        public virtual double Height
        {
            get
            {
                return Size.Y;
            }
            set
            {
                Size = new Vector2(Size.X, value);
            }
        }
        internal double GlobalLeft
        {
            get
            {
                return Position.X;
            }
        }
        internal double GlobalTop
        {
            get
            {
                return Position.Y;
            }
        }

        //Quick-access to underlying transform

        ///// <summary>
        ///// Local position X
        ///// X-position from parent transform position
        ///// </summary>
        //internal override double X { get { return Left; } set { Left = value; } }

        ///// <summary>
        ///// Local position Y
        ///// Y-position from parent transform position
        ///// </summary>
        //internal override double Y { get { return Top; } set { Top = value; } }

        //Gets global position x and y
        public double GlobalX { get { return Position.X; } }
        public double GlobalY { get { return Position.Y; } }

        public void SetLeft(double x)
        {
            Left = x;
        }
        public double GetLeft()
        {
            return Left;
        }
    }
}
