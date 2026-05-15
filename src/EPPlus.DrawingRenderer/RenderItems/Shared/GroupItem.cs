using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Collections.Generic;

namespace EPPlus.DrawingRenderer.RenderItems
{
    internal abstract class GroupItem : RenderItem
    {
        /// <summary>
        /// In degrees
        /// </summary>
        internal double Rotation = double.NaN;
        /// <summary>
        /// The translated position of this item in points
        /// Also the parent position of the group item 
        /// (This may seem strange but it ensures the the translation is seen 
        /// immediately in the global position of GroupItem without affecting local position)
        /// </summary>
        internal Point TranslationOffset = new Point(0,0);

        Point _altRotationPoint = null;

        internal Point RotationPoint 
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

        internal Coordinate Scale = null;


        //Transform _rotationPoint;

        /// <summary>
        /// Items contained in this group
        /// </summary>
        internal protected List<RenderItem> _childItems = new List<RenderItem>();

        public GroupItem(DrawingRenderer renderer) : base(renderer)
        {
            Bounds.Parent = TranslationOffset;
            //_rotationPoint = Bounds;
        }

        public GroupItem(DrawingRenderer renderer, double localXPos, double localYPos) : this(renderer)
        {
            Bounds.Left = localXPos;
            Bounds.Top = localYPos;
        }


        public GroupItem(DrawingBase renderer, BoundingBox parent, double rotation, Transform rotationPoint = null) : this(renderer, 0, 0)
        {
            TranslationOffset.Parent = parent;
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
