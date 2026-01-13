using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBodyItem : TextBody<SvgParagraphItem>
    {
        public SvgTextBodyItem(BoundingBox parent) : base(parent)
        {
        }

        internal override List<SvgParagraphItem> Paragraphs { get; set; } = new List<SvgParagraphItem>();

        public override void Render(StringBuilder sb)
        {
            sb.AppendLine($"<g transform=\"translate({Bounds.GlobalX},{Bounds.GlobalY})\" ");
            base.Render(sb);
            sb.Append(" >");
            sb.AppendLine($"<title>txtBody</title>");

            foreach(var item in Paragraphs)
            {
                item.Render(sb);
            }
            sb.AppendLine("</g>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }
    }
}
