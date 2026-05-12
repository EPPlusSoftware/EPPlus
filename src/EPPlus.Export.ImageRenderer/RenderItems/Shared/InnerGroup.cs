using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class InnerGroup : RenderItem
    {
        /// <summary>
        /// Items contained in this group
        /// </summary>
        internal protected List<RenderItem> _childItems = new List<RenderItem>();

        public InnerGroup(DrawingBase renderer) : base(renderer)
        {

        }

        //public InnerGroup(DrawingBase renderer, double localXPos, double localYPos) : this(renderer)
        //{
        //    Bounds.Left = localXPos;
        //    Bounds.Top = localYPos;
        //}


        //public InnerGroup(DrawingBase renderer, BoundingBox parent, double rotation, Transform rotationPoint = null) : this(renderer, 0, 0)
        //{
        //    TranslationOffset.Parent = parent;
        //    Rotation = rotation;
        //    if (rotationPoint != null)
        //    {
        //        RotationPoint = new Point(rotationPoint.LocalPosition.X, rotationPoint.LocalPosition.Y);
        //    }
        //}

        //internal void SetRotationPointToCenterOfGroup(double rotation = double.NaN)
        //{
        //    RotationPoint = new Point(Bounds.Width / 2, Bounds.Height / 2);

        //    if (double.IsNaN(rotation) == false)
        //    {
        //        Rotation = rotation;
        //    }
        //}

        internal void AddChildItem(RenderItem item)
        {
             item.Bounds.Parent = Bounds;

            _childItems.Add(item);

            Bounds.Width = item.Bounds.Right > Bounds.Width ? item.Bounds.Right : Bounds.Width;
            Bounds.Height = item.Bounds.Bottom > Bounds.Height ? item.Bounds.Bottom : Bounds.Height;
        }

        public override RenderItemType Type => RenderItemType.Group;
    }
}
