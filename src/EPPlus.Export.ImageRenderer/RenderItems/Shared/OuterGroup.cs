using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class OuterGroup : RenderItem
    {
        InnerGroup _innerItems;

        /// <summary>
        /// In degrees
        /// </summary>
        internal double Rotation = double.NaN;

        BoundingBox _altRotationPoint = null;

        internal BoundingBox RotationPoint
        {
            get
            {
                if (_altRotationPoint == null)
                {
                    return Bounds;
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

        public OuterGroup(DrawingBase renderer) : base(renderer)
        {
            _innerItems.Bounds.Parent = Bounds;
            //Bounds.Parent = TranslationOffset;
            //_rotationPoint = Bounds;
        }

        public OuterGroup(DrawingBase renderer, double localXPos, double localYPos) : this(renderer)
        {
            Bounds.Left = localXPos;
            Bounds.Top = localYPos;
        }


        public OuterGroup(DrawingBase renderer, BoundingBox parent, double rotation, Transform rotationPoint = null) : this(renderer, 0, 0)
        {
            Bounds.Parent = parent;
            Rotation = rotation;
            if (rotationPoint != null)
            {
                RotationPoint = new BoundingBox(rotationPoint.LocalPosition.X, rotationPoint.LocalPosition.Y);
            }
        }

        internal void SetRotationPointToCenterOfGroup(double rotation = double.NaN)
        {
            RotationPoint = new BoundingBox(Bounds.Width / 2, Bounds.Height / 2);

            if (double.IsNaN(rotation) == false)
            {
                Rotation = rotation;
            }
        }

        internal void AddChildItem(RenderItem item)
        {
            if (item is GroupItem)
            {
                var subGroup = (GroupItem)item;
                subGroup.TranslationOffset.Parent = Bounds;
            }
            else
            {
                item.Bounds.Parent = Bounds;
            }
            _childItems.Add(item);

            Bounds.Width = item.Bounds.Right > Bounds.Width ? item.Bounds.Right : Bounds.Width;
            Bounds.Height = item.Bounds.Bottom > Bounds.Height ? item.Bounds.Bottom : Bounds.Height;
        }

        public override RenderItemType Type => RenderItemType.Group;
    }
}
