using EPPlus.Export.ImageRenderer.RenderItems.Independent.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.SvgItem
{
    internal class SvgIndependentCanvas : IndependentCanvas
    {
        List<RenderItemIndependent> renderItems = new List<RenderItemIndependent>();

        public SvgIndependentCanvas(BoundingBox canvasBounds, Color bgColor) : base(canvasBounds, bgColor)
        {
        }

        public override RenderItemType Type => RenderItemType.Rect;

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<svg width=\"{Bounds.Width.PointToPixel()}\" height=\"{Bounds.Height.PointToPixel()}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"default\" Overflow=\"Hidden\">");
            Background.Render(sb);
            foreach (var item in renderItems)
            {
                item.Render(sb);
            }
            sb.Append($"</svg>");
        }

        internal void AddRenderItem(RenderItemIndependent renderItem)
        {
            renderItems.Add(renderItem);
        }

        internal override IndependentRect CreateBackground(BoundingBox bounds)
        {
            var bg = new SvgIndependentRect(Bounds, bounds);
            return bg;
        }
    }
}
