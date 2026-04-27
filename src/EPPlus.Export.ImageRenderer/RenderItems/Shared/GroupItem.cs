using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class GroupItem : RenderItem
    {
        /// <summary>
        /// In degrees
        /// </summary>
        internal double Rotation = double.NaN;
        /// <summary>
        /// In points
        /// </summary>
        internal Point Position = null;

        Point _altRotationPoint = null;

        internal Point RotationPoint 
        {
            get
            {
                if (_altRotationPoint == null)
                {
                    return Position;
                }
                return _altRotationPoint;
            }
            set
            {
                _altRotationPoint = value;
            }
        }

        internal Coordinate Scale = null;


        //Transform _rotationPoint;

        /// <summary>
        /// Items contained in this group
        /// </summary>
        internal protected List<RenderItem> _childItems = new List<RenderItem>();

        public GroupItem(DrawingBase renderer) : base(renderer)
        {
            Bounds.Parent = Position;
            //_rotationPoint = Bounds;
        }

        public GroupItem(DrawingBase renderer, double localXPos, double localYPos) : this(renderer)
        {
            Position = new Point(localXPos, localYPos);
        }


        public GroupItem(DrawingBase renderer, BoundingBox parent, double rotation, Transform rotationPoint = null) : this(renderer, 0, 0)
        {
            Position.Parent = parent;
            Rotation = rotation;
            if (rotationPoint != null)
            {
                RotationPoint = new Point(rotationPoint.LocalPosition.X, rotationPoint.LocalPosition.Y);
            }
        }

        internal void SetRotationPointToCenterOfGroup(double rotation = double.NaN)
        {
            RotationPoint = new Point(Bounds.Width/2, Bounds.Height/2);

            if (double.IsNaN(rotation) == false)
            {
                Rotation = rotation;
            }
        }

        internal void AddChildItem(RenderItem item)
        {
            item.Bounds.Parent = Position;
            _childItems.Add(item);

            Bounds.Width = item.Bounds.Right > Bounds.Width ? item.Bounds.Right : Bounds.Width;
            Bounds.Height = item.Bounds.Bottom > Bounds.Height ? item.Bounds.Bottom : Bounds.Height;
        }

        public override RenderItemType Type => RenderItemType.Group;
    }
}
