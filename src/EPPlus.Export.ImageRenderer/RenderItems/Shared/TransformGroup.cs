using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class TransformGroup : RenderItem
    {
        protected InnerGroup _innerGroup;

        Point PositionAfterTransform;

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
        Coordinate _scale = new Coordinate(1,1);

        internal Coordinate Scale
        {
            get 
            {
                return _scale;
            } 
            set
            { 
                Bounds.Parent.Scale = new Graphics.Math.Vector2(value.X, value.Y); 
                _scale = value;
            } 
        }

        public TransformGroup(DrawingBase renderer) : base(renderer)
        {
            _innerGroup = CreateInnerGroup();
            _innerGroup.Bounds.Parent = Bounds;
        }

        public TransformGroup(DrawingBase renderer, double localXPos, double localYPos) : this(renderer)
        {
            Bounds.Left = localXPos;
            Bounds.Top = localYPos;
        }


        public TransformGroup(DrawingBase renderer, BoundingBox parent, double rotation, Transform rotationPoint = null) : this(renderer, 0, 0)
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

        /// <summary>
        /// Adds child item to the group under this group
        /// </summary>
        /// <param name="item"></param>
        internal void AddChildItem(RenderItem item)
        {
            _innerGroup.AddChildItem(item);

            Bounds.Width = item.Bounds.Right > Bounds.Width ? item.Bounds.Right : Bounds.Width;
            Bounds.Height = item.Bounds.Bottom > Bounds.Height ? item.Bounds.Bottom : Bounds.Height;
        }

        internal abstract InnerGroup CreateInnerGroup();

        public override RenderItemType Type => RenderItemType.Group;
    }
}
