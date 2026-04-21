using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgContainerItem : ContainerItem
    {
        public SvgContainerItem(RenderItem innerItem, RenderItem outerItem) : base(innerItem, outerItem)
        {
        }

        public override void Render(StringBuilder sb)
        {
            var grpItem = new SvgGroupItem(DrawingRenderer, Bounds.Left, Bounds.Top);
            grpItem.Render(sb);

            OuterItem.Render(sb);

            var grpItem2 = new SvgGroupItem(DrawingRenderer, MarginLeft, MarginTop);
            grpItem2.Render(sb);

            InnerItem.Render(sb);

            var endGroupItem2 = new SvgEndGroupItem(DrawingRenderer, Bounds);
            endGroupItem2.Render(sb);


            var endGroupItem = new SvgEndGroupItem(DrawingRenderer, Bounds);
            endGroupItem.Render(sb);
        }
    }
}
