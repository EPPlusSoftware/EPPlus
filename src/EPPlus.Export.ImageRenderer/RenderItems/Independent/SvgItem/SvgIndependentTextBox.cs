using EPPlus.Export.ImageRenderer.RenderItems.Independent.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.SvgItem
{
    internal class SvgIndependentTextBox : IndependentTextBox
    {

        public SvgIndependentTextBox(DrawingBase baseObj, BoundingBox bounds) : base(baseObj, bounds)
        {
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            TextBody.AppendRenderItems(renderItems);
        }

        internal override TextBodyItem CreateTextBody(DrawingBase obj, BoundingBox parent)
        {
            return new SvgTextBodyItem(obj, parent, true);
        }
    }
}
