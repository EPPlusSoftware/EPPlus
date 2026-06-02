using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;


namespace EPPlus.DrawingRenderer.RenderItems.SvgItem
{
    public class SvgTextBodyRenderItem : RenderTextBody
    {
        public SvgTextBodyRenderItem(BoundingBox parent, bool autoSize) : base(parent, autoSize)
        {
        }

        public SvgTextBodyRenderItem(BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize = false) : base(parent, left, top, maxWidth, maxHeight, clampedToParent, autoSize)
        {

        }

        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, string textIfEmpty = "")
        {
            return new SvgParagraphRenderItem(this, parent, textIfEmpty);
        }

        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, IRichTextFormatSimple richText)
        {
            return new SvgParagraphRenderItem(this, parent, richText);
        }
    }
}
