using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer.RenderItems.SvgItem
{
    public class SvgParagraphRenderItem : ParagraphRenderItem
    {
        public SvgParagraphRenderItem(RenderTextBody body, BoundingBox parent) : base(parent)
        {

        }

        public override RenderItemType Type => RenderItemType.Paragraph;
    }
}
