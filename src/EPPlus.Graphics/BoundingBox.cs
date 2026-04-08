using EPPlus.Graphics.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    internal class BoundingBox : Transform
    {
        internal BoundingBox() : base()
        {
        }

        internal BoundingBox(double width, double height) : base(0, 0, width, height)
        {
        }
        internal BoundingBox(double left, double top, double width, double height) : base(left, top, width, height)
        {
        }

        /// <summary>
        /// Y pos (min)
        /// </summary>
        internal double Top
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
        internal double Left
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
        internal double Bottom
        {
            get
            {
                return LocalPosition.Y + Size.Y;
            }
        }

        /// <summary>
        /// If @ClampedToParent is true will not set value beyond parent
        /// </summary>
        internal double Right
        {
            get
            {
                return LocalPosition.X + Size.X;
            }
        }
        internal virtual double Width
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

        internal virtual double Height
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
        internal double GlobalX { get { return Position.X; } }
        internal double GlobalY { get { return Position.Y; } }
    }
}
