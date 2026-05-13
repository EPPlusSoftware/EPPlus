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
