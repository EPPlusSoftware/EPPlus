using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using System.Drawing;


namespace EPPlus.DrawingRenderer.RenderItems.SvgItem
{
    public class SvgTextBodyRenderItem : RenderTextBody
    {
        public SvgTextBodyRenderItem(RenderContext renderContext, BoundingBox parent, bool autoSize) : base(renderContext, parent, autoSize)
        {
        }

        public SvgTextBodyRenderItem(RenderContext renderContext, BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize = false) : base(renderContext, parent, left, top, maxWidth, maxHeight, clampedToParent, autoSize)
        {

        }

        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, string textIfEmpty = "")
        {
            return new SvgParagraphRenderItem(RenderContext, this, parent, textIfEmpty);
        }

        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, IRichTextFormatSimple richText)
        {
            return new SvgParagraphRenderItem(RenderContext, this, parent, richText);
        }
    }
}
