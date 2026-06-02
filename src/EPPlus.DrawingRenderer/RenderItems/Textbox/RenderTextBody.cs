using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;

namespace EPPlus.DrawingRenderer.RenderItems
{
    /// <summary>
    /// Text anchoring
    /// </summary>
    public enum TextAnchoringType
    {
        /// <summary>
        /// Anchor the text to the bottom
        /// </summary>
        Bottom,
        /// <summary>
        /// Anchor the text to the center
        /// </summary>
        Center,
        /// <summary>
        /// Anchor the text so that it is distributed vertically.
        /// </summary>
        Distributed,
        /// <summary>
        /// Anchor the text so that it is justified vertically.
        /// </summary>
        Justify,
        /// <summary>
        /// Anchor the text to the top
        /// </summary>
        Top
    }

    public abstract class RenderTextBody : GroupRenderItem
    {
        public RenderTextBody(BoundingBox parent, bool autoSize)
        {
            Bounds.Parent = parent;
            AutoSize = autoSize;
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
            Bounds.Name = "Textbody";
        }
        public RenderTextBody(BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize=false) : this(parent, autoSize)
        {
            Bounds.Left = left;
            Bounds.Top = top;
            Bounds.Width = maxWidth;
            Bounds.Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
            Bounds.Name = "Textbody";
        }

        public List<ParagraphRenderItem> Paragraphs { get; set; } = new List<ParagraphRenderItem>();

        public TextAnchoringType VerticalAlignment = TextAnchoringType.Top;
        public string Text { get; set; }
        public double MaxWidth { get; set; }
        public double MaxHeight { get; set; }
        /// <summary>
        /// Shorthand for Bounds.Width
        /// </summary>
        public double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }

        /// <summary>
        /// Shorthand for Bounds.Height
        /// </summary>
        public double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }

        public bool AutoSize { get; set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public double RightMargin { get; set; }
        public double LeftMargin { get; set; }
        public string FontColorString { get; set; }

        
        public void AppendRenderItems(List<RenderItem> renderItems)
        {
            //foreach(var item in Paragraphs)
            //{
            //    AddChildItem(item);
            //}
            //GroupRenderItem groupItem;
            //if (Bounds.Parent.Rotation == 0) //If the parent is rotated, we should not apply rotation again. This is usually when the parent is a textbox.
            //{
            //    groupItem = new GroupRenderItem(Bounds, Bounds.Rotation);
            //}
            //else
            //{
            //    groupItem = new GroupRenderItem(Bounds);
            //}

            //if (FontColorString != null)
            //{
            //    groupItem.GroupTransform += $" fill=\"{FontColorString}\"";
            //}
            //renderItems.Add(groupItem);

            //Set bounds position to be translation
            //Posibly remove translationOffset and make it always be bounds?
            //But then we will have an inaccurate bounding box if a child object has negative position.
            //TranslationOffset.Left = Bounds.Left;
            //TranslationOffset.Top = Bounds.Top;

            renderItems.Add(this);
            foreach (var item in Paragraphs)
            {
                AddChildItem(item);
            }
        }

        public void AddParagraph(string text = null)
        {
            double paragraphTop = 0;

            if(Paragraphs.Count != 0)
            {
                paragraphTop = Paragraphs.Last().Bounds.Bottom;
            }
            var paragraph = CreateParagraph(Bounds, text);
            paragraph.Bounds.Name = $"Container{Paragraphs.Count}";
            paragraph.Bounds.Top = paragraphTop;
            Text = text;

            if (AutoSize)
            {
                if (Paragraphs.Count == 0)
                {
                    Bounds.Height = paragraph.Bounds.Height;
                }
                else
                {
                    Bounds.Height += paragraph.Bounds.Height;
                }

                if (Bounds.Width < paragraph.Bounds.Width || (Bounds.Width == MaxWidth && Paragraphs.Count == 0))
                {
                    Bounds.Width = paragraph.Bounds.Width;
                }
            }
            Paragraphs.Add(paragraph);
        }

        /// <summary>
        /// Get the start of text space vertically
        /// </summary>
        /// <returns></returns>
        protected double GetAlignmentVertical()
        {
            double alignmentY = 0;

            switch (VerticalAlignment)
            {
                case TextAnchoringType.Top:
                    alignmentY = Bounds.Top;
                    break;
                //Center means center of a Shape's ENTIRE bounding box height.
                //Not center of the Inset GetRectangle
                case TextAnchoringType.Center:
                    alignmentY = (MaxHeight - Bounds.Height) / 2 + Bounds.Top;
                    break;
                case TextAnchoringType.Bottom:
                    alignmentY = MaxHeight - Bounds.Height;
                    break;
            }

            return alignmentY;
        }

        protected abstract ParagraphRenderItem CreateParagraph(BoundingBox parent, string textIfEmpty = "");
    }
}
