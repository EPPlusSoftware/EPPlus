using EPPlus.Export.ImageRenderer.RenderItems.Independent.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.SvgItem
{
    internal class SvgIndependentRect : IndependentRect
    {
        public SvgIndependentRect(BoundingBox parent) : base(parent)
        {
        }

        public SvgIndependentRect(BoundingBox parent, BoundingBox bounds) : base(parent, bounds)
        {
        }

        public SvgIndependentRect(BoundingBox parent, double width, double height) : base(parent)
        {
            Bounds.Width = width;
            Bounds.Height = height;
        }

        public override void Render(StringBuilder sb)
        {
            //var groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
            //groupItem.Render(sb);

            RenderRect(sb);

            //groupItem.RenderEndGroup(sb);
        }

        internal void RenderRect(StringBuilder sb)
        {
            sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
                Left.PointToPixelString(),
                Top.PointToPixelString(),
                Width.PointToPixelString(),
                Height.PointToPixelString());
            SvgBaseRendererIndependent.BaseRender(sb, this);
            sb.AppendFormat("/>");
        }
    }
}
