using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;

namespace EPPlus.DrawingRenderer.RenderItems
{
    public class RenderTextBody
    {
        public RenderTextBody(BoundingBox parent, bool autoSize)
        {
            Bounds.Parent = parent;
            AutoSize = autoSize;
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
        }
        public RenderTextBody(BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize=false) : this(parent, autoSize)
        {
            Bounds.Left = left;
            Bounds.Top = top;
            Bounds.Width = maxWidth;
            Bounds.Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
        }

        public List<RenderParagraph> Paragraphs { get; set; } = new List<RenderParagraph>();
        public BoundingBox Bounds { get; private set; } = new BoundingBox();
        public double MaxWidth { get; set; }
        public double MaxHeight { get; set; }
        public bool AutoSize { get; private set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public double RightMargin { get; set; }
        public double LeftMargin { get; set; }

        //internal override void AppendRenderItems(List<RenderItem> renderItems)
        //{
        //    SvgGroupItem groupItem;
        //    if (Bounds.Parent.Rotation == 0) //If the parent is rotated, we should not apply rotation again. This is usually when the parent is a textbox.
        //    {
        //        groupItem = new SvgGroupItem(DrawingRenderer, Bounds, Bounds.Rotation);
        //    }
        //    else
        //    {
        //        groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
        //    }

        //    if (FontColorString != null)
        //    {
        //        groupItem.GroupTransform += $" fill=\"{FontColorString}\"";
        //    }

        //    renderItems.Add(groupItem);
        //    foreach (SvgParagraphItem item in Paragraphs)
        //    {
        //        renderItems.Add(item);
        //    }
        //    renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        //}

        //internal override Paragraph CreateParagraph(TextBodyItem textBody, BoundingBox parent)
        //{
        //    return new SvgParagraphItem(this, DrawingRenderer, parent);
        //}

        //internal override Paragraph CreateParagraph(TextBodyItem textBody, ExcelDrawingParagraph paragraph, BoundingBox parent, string textIfEmpty = null)
        //{
        //    return new SvgParagraphItem(this, DrawingRenderer, parent, paragraph, textIfEmpty);
        //}

        //internal override Paragraph CreateParagraph(TextBodyItem textBody, BoundingBox parent, string textIfEmpty = "")
        //{
        //    return new SvgParagraphItem(this, DrawingRenderer, parent, textIfEmpty);
        //}
    }
}
