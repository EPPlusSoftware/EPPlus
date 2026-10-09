using EPPlus.Graphics.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

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
        public virtual double Top
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
        public virtual double Left
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

        public string TextAnchor { get; set; }
        //public double Rotation { get; set; }
        //public string GroupTransform = "";
        //public List<RenderItem> RenderItems { get; } = new List<RenderItem>();

        Graphics.TranformPoint _altRotationPoint = null;
        /// <summary>
        /// The translated position of this item in points
        /// Also the parent position of the group item 
        /// (This may seem strange but it ensures the the translation is seen 
        /// immediately in the global position of GroupItem without affecting local position)
        /// </summary>
        public Graphics.TranformPoint TranslationOffset = new Graphics.TranformPoint(0, 0);
        public Graphics.TranformPoint RotationPoint
        {
            get
            {
                if (_altRotationPoint == null)
                {
                    return TranslationOffset;
                }
                return _altRotationPoint;
            }
            set
            {
                _altRotationPoint = value;
            }
        }

        //public Coordinate Scale = null;

        internal void SetRotationPointToCenterOfGroup(double rotation = double.NaN)
        {
            RotationPoint = new Graphics.TranformPoint(Width / 2, Height / 2);

            if (double.IsNaN(rotation) == false)
            {
                Rotation = rotation;
            }
        }
        public double GroupWidth
        {
            get
            {
                var left = Left;
                var right = Right;
                foreach (var item in ChildObjects)
                {
                    if (item is BoundingBox bb)
                    {
                        left = bb.Left > left ? bb.Left : left;
                        right = bb.Right > right ? bb.Right : right;
                    }
                }
                return right - left;
            }
        }

        //Note: This does not take negative child items into acount
        //TODO: Fix that
        public double GroupHeight
        {
            get
            {
                var top = Top;
                var bottom = Bottom;
                foreach (var item in ChildObjects)
                {
                    if (item is BoundingBox bb)
                    {
                        top = bb.Top > top ? bb.Top : top;
                        bottom = bb.Bottom > bottom ? bb.Bottom : bottom;
                    }
                }
                return bottom - top;
            }
        }

        public string UniqueKey 
        { 
            get 
            { 
                return $"{Left} {Top} {Width} {Height}";
            } 
        }

        //We've got a huge problem.
        //Rotation breaks our terms of "top" and "left" as parental points may have rotated the item meaning we may return a "top" that is on the global bottom of the item
    }
}
