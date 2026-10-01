using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
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
        public RenderTextBody(RenderContext renderContext, BoundingBox parent, bool autoSize)
        {
            RenderContext = renderContext;
            Parent = parent;
            AutoSize = autoSize;
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
            Name = "Textbody";
        }
        public RenderTextBody(RenderContext renderContext, BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize=false) : this(renderContext, parent, autoSize)
        {
            RenderContext = renderContext;
            Left = left;
            Top = top;
            Width = maxWidth;
            Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
            Name = "Textbody";
        }

        protected RenderContext RenderContext { get; private set; }
        public List<ParagraphRenderItem> Paragraphs { get; set; } = new List<ParagraphRenderItem>();

        public TextAnchoringType VerticalAlignment = TextAnchoringType.Top;
        public string Text { get; set; }
        public double MaxWidth { get; set; }
        public double MaxHeight { get; set; }
        /// <summary>
        /// Shorthand for Width
        /// </summary>
        public double Width { get { return Width; } set { Width = value; } }

        /// <summary>
        /// Shorthand for Height
        /// </summary>
        public double Height { get { return Height; } set { Height = value; } }

        public bool AutoSize { get; set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }
        public double RightMargin { get; set; }
        public double LeftMargin { get; set; }
        public string FontColorString { get; set; }

        
        public void AppendRenderItems(List<Transform> renderItems)
        {
            //foreach(var item in Paragraphs)
            //{
            //    AddChildItem(item);
            //}
            //GroupRenderItem groupItem;
            //if (Parent.Rotation == 0) //If the parent is rotated, we should not apply rotation again. This is usually when the parent is a textbox.
            //{
            //    groupItem = new GroupRenderItem(Bounds, Rotation);
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
            //TranslationOffset.Left = Left;
            //TranslationOffset.Top = Top;

            renderItems.Add(this);

            var titleItem = new TitleRenderItem("TextBody group");
            AddChildItem(titleItem);
            foreach (var item in Paragraphs)
            {
                AddChildItem(item);
            }
        }

        public ParagraphRenderItem AddParagraph(IRichTextFormatSimple rtFormat)
        {
            var paragraph = CreateParagraph(this, rtFormat);
            AdjustAndAddParagraph(paragraph);
            return paragraph;
        }

        public ParagraphRenderItem AddParagraph(string text = null)
        {
            var paragraph = CreateParagraph(this, text);
            AdjustAndAddParagraph(paragraph);
            return paragraph;
        }

        public void ApplyAutoSize()
        {
            if (AutoSize)
            {
               var currentHeight = 0d;
               var currentWidth = 0d;

                foreach(var paragraph in Paragraphs)
                {
                    currentHeight += paragraph.Height;

                    if (currentWidth < paragraph.Width || currentWidth == MaxWidth)
                    {
                        currentWidth = paragraph.Width;
                    }
                }

                Width = currentWidth;
                Height = currentHeight;
            }
        }

        /// <summary>
        /// If text is added to the first paragraph without using textbody e.g. Paragraphs[0].AddText()
        /// Subsequent paragraphs must be updated
        /// </summary>
        public void RecalculateParagraphs()
        {
            if(Paragraphs != null && Paragraphs.Count != 0)
            {
                double lastParagraphBottom = Paragraphs[0].Top;

                double smallestLeft = double.MaxValue;
                double largestWidth = double.MinValue;
                double totalHeight = 0;

                foreach (var paragraph in Paragraphs)
                {
                    paragraph.Top = lastParagraphBottom;
                    lastParagraphBottom = paragraph.Bottom;

                    smallestLeft = Math.Min(smallestLeft, paragraph.Left);
                    largestWidth = Math.Max(largestWidth, paragraph.Width);
                    totalHeight += paragraph.Height;
                }

                ContentBounds.Top = Paragraphs[0].Top;
                ContentBounds.Left = smallestLeft;
                ContentBounds.Width = largestWidth;
                ContentBounds.Height = totalHeight;

                if (AutoSize)
                {
                    Height = totalHeight;
                    Width = ContentBounds.Width;
                }
            }
        }

        /// <summary>
        /// The total bounds of all paragraphs without margins
        /// </summary>
        protected BoundingBox ContentBounds = new BoundingBox();

        private void AdjustAndAddParagraph(ParagraphRenderItem paragraph)
        {
            paragraph.Name = $"Container{Paragraphs.Count}";
            paragraph.Top = GetTopToAddNextParagraphAt();

            if (AutoSize)
            {
                if (Paragraphs.Count == 0)
                {
                    Height = paragraph.Height;
                }
                else
                {
                    Height += paragraph.Height;
                }

                if (Width < paragraph.Width || (Width == MaxWidth && Paragraphs.Count == 0))
                {
                    Width = paragraph.Width;
                }
            }
            Paragraphs.Add(paragraph);
            RecalculateParagraphs();
        }

        private double GetTopToAddNextParagraphAt()
        {
            double paragraphTop = 0;

            if (Paragraphs.Count != 0)
            {
                paragraphTop = Paragraphs.Last().Bottom;
            }
            return paragraphTop;
        }


        /// <summary>
        /// Get the start of text space vertically
        /// </summary>
        /// <returns></returns>
        public double GetAlignmentVertical()
        {
            double alignmentY = 0;

            switch (VerticalAlignment)
            {
                case TextAnchoringType.Top:
                    alignmentY = Top;
                    break;
                //Center means center of a Shape's ENTIRE bounding box height.
                //Not center of the Inset GetRectangle
                case TextAnchoringType.Center:
                    if(AutoSize == false)
                    {
                        alignmentY = (Height - ContentBounds.Height) / 2d;
                    }
                    break;
                case TextAnchoringType.Bottom:
                    alignmentY = Height - ContentBounds.Height;
                    break;
            }

            return alignmentY;
        }

        protected abstract ParagraphRenderItem CreateParagraph(BoundingBox parent, string textIfEmpty = "");

        protected abstract ParagraphRenderItem CreateParagraph(BoundingBox parent, IRichTextFormatSimple richText);
    }
}
