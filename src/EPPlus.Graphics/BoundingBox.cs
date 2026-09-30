using EPPlus.Graphics.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    public class BoundingBox<T> : Transform<T> where T : Transform<T>
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
        public double GlobalLeft
        {
            get
            {
                return Position.X;
            }
        }
        public double GlobalTop
        {
            get
            {
                return Position.Y;
            }
        }
        public double GlobalRight => GlobalLeft + Width;
        public double GlobalBottom => GlobalTop + Height;

        public string UniqueKey 
        { 
            get 
            { 
                return $"{Left} {Top} {Width} {Height}";
            } 
        }
    }
}
