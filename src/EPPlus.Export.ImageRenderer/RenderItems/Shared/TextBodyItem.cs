using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;


namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    /// <summary>
    /// Margin left = X
    /// Margin right = Y
    /// </summary>
    internal abstract class TextBodyItem : DrawingObject
    {
        /// <summary>
        /// Shorthand for Bounds.Width
        /// </summary>
        internal double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }

        /// <summary>
        /// Shorthand for Bounds.Height
        /// </summary>
        internal double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }

        internal string FontColorString { get; set; }

        internal eTextAnchoringType VerticalAlignment = eTextAnchoringType.Top;

        internal abstract List<ParagraphItem> Paragraphs { get; set; }

        public bool AllowOverflow;

        internal bool WrapText = true;

        private FontMeasurerTrueType _measurer = null;
        internal string _text;
        public TextBodyItem(DrawingBase renderer, BoundingBox parent)  : base(renderer, parent)
        {
            Bounds.Name = "TxtBody";

            Bounds.Parent = parent;

            Bounds.Width = parent.Width;
            Bounds.Height = parent.Height;
        }

        public void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text=null)
        {
            var measureFont = item.DefaultRunProperties.GetMeasureFont();
            bool isFirst = Paragraphs.Count == 0;

            var paragraph = CreateParagraph(item, Bounds);
            paragraph.Bounds.Name = $"Container{Paragraphs.Count}";
            if (string.IsNullOrEmpty(text) == false)
            {
                paragraph.AddText(text, item.DefaultRunProperties);
            }
            paragraph.Bounds.Top = startingY;
            _text = text;
            Paragraphs.Add(paragraph);
        }

        internal virtual void ImportTextBody(ExcelTextBody body)
        {
            _text = null;
            VerticalAlignment = body.Anchor;
            //We already apply bounds top via the parent Transform
            double paragraphStartY = GetAlignmentVertical();

            foreach (var paragraph in body.Paragraphs)
            {
                ImportParagraph(paragraph, paragraphStartY);
                var addedPara = Paragraphs.Last();
                paragraphStartY = addedPara.Bounds.Bottom;
            }
            if (Paragraphs != null && Paragraphs.Count() > 0)
            {
                Bounds.Height = paragraphStartY;
            }
        }

        //public void AddText(string text, FontMeasurerTrueType measurer)
        //{
        //    if (Paragraphs.Count == 0)
        //    {
        //        AddParagraph(text, measurer);
        //    }
        //    else
        //    {
        //        Paragraphs.Last().AddText(text, measurer);
        //    }
        //}
        //internal void AddText(string text, ExcelTextFont font)
        //{
        //    //Document Top position for the paragraph text based on vertical alignment
        //    var posY = GetAlignmentVertical();
        //    //var vertAlignAttribute = GetVerticalAlignAttribute(posY);

        //    //var measurer = font.PictureRelationDocument.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
        //    //Limit bounding area with the space taken by previous paragraphs
        //    //Note that this is ONLY identical to PosY if the vertical alignment is top

        //    //The first run in the first paragraph must apply different line-spacing
        //    //var svgParagraph = new SvgParagraph(text, font, area, vertAlignAttribute, posY);
        //    var paragraph = CreateParagraph(Bounds);
        //    paragraph.AddText(text, font);
        //    paragraph.FillColor = font.Fill.Color.To6CharHexString();

        //    Paragraphs.Add(paragraph);
        //}
        public void SetMeasurer(FontMeasurerTrueType fontMeasurer)
        {
            _measurer = fontMeasurer;
        }

        double? _alignmentY = null;

        /// <summary>
        /// Get the start of text space vertically
        /// </summary>
        /// <param name="fontSizeInPixels"></param>
        /// <returns></returns>
        private double GetAlignmentVertical()
        {
            double alignmentY = 0;

            switch (VerticalAlignment)
            {
                case eTextAnchoringType.Top:
                    alignmentY = 0;
                    break;
                case eTextAnchoringType.Center:
                    var adjustedHeight = ((Bounds.Height - Bounds.Parent.LocalPosition.Y) / 2);

                    alignmentY = adjustedHeight;
                    break;
                case eTextAnchoringType.Bottom:
                    alignmentY = Bounds.Height;
                    break;
            }

            _alignmentY = alignmentY;

            return _alignmentY.Value;
        }


        /// <summary>
        /// Each file format defines its own paragraph
        /// </summary>
        /// <returns></returns>
        internal abstract ParagraphItem CreateParagraph(ExcelDrawingParagraph paragraph, BoundingBox parent);

        /// <summary>
        /// Each file format defines its own paragraph
        /// </summary>
        /// <returns></returns>
        internal abstract ParagraphItem CreateParagraph(BoundingBox parent);
    }
}
