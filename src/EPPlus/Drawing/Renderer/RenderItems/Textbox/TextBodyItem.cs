using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;


namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    /// <summary>
    /// Margin left = X
    /// Margin right = Y
    /// </summary>
    internal abstract class TextBodyItem
    {
        public TextBodyItem(DrawingBase renderer, BoundingBox parent, bool autoSize) : base(renderer, parent)
        {
            Bounds.Name = "TxtBody";
            Bounds.Parent = parent;
            AutoSize = autoSize;
        }
        public TextBodyItem(DrawingBase renderer, BoundingBox parent, double maxWidth, double maxHeight) : base(renderer, parent)
        {
            Bounds.Name = "TxtBody";
            Bounds.Parent = parent;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
            AutoSize = false;
        }
        internal bool AutoSize { get; set; }
        internal double MaxWidth { get; set; }
        internal double MaxHeight { get; set; }

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

        internal string _text;

        public void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text=null)
        {
            var measureFont = item.DefaultRunProperties.GetMeasureFont();
            bool isFirst = Paragraphs.Count == 0;

            var paragraph = CreateParagraph(this, item, Bounds, text);
            paragraph.Bounds.Name = $"Container{Paragraphs.Count}";
            paragraph.Bounds.Top = startingY;
            _text = text;

            if(AutoSize)
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

        public void AddParagraph(double startingY, string text = null)
        {
            var paragraph = CreateParagraph(this, Bounds, text);
            paragraph.Bounds.Name = $"Container{Paragraphs.Count}";
            paragraph.Bounds.Top = startingY;
            _text = text;

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

        internal void SetHorizontalAlignmentPosition()
        {
            //if (AutoSize)
            //{
                foreach (var p in Paragraphs)
                {
                    switch (p.HorizontalAlignment)
                    {
                        case eTextAlignment.Left:
                            p.Bounds.Left = 0;
                            break;
                        case eTextAlignment.Center:
                            p.Bounds.Left = (Bounds.Width / 2) - (p.Bounds.Width / 2);
                            break;
                        case eTextAlignment.Right:
                            p.Bounds.Left = Bounds.Right - p.Bounds.Width;
                            break;
                        case eTextAlignment.Distributed:
                        case eTextAlignment.Justified:
                        case eTextAlignment.JustifiedLow:
                        case eTextAlignment.ThaiDistributed:
                            p.Bounds.Left = 0;                    //TODO: Set left for now as we do not support distributed spacing yet
                            break;
                    }
                }
            //}
        }

        internal virtual void ImportTextBody(ExcelTextBody body, ExcelHorizontalAlignment horizontalDefault = ExcelHorizontalAlignment.Left)
        {
            _text = null;
            VerticalAlignment = body.Anchor;

            //We already apply bounds top via the parent Transform
            double currentHeight = 0;
            double largestWidth = double.MinValue;

            foreach (var paragraph in body.Paragraphs)
            {
                ImportParagraph(paragraph, currentHeight);
                var addedPara = Paragraphs.Last();
                currentHeight = addedPara.Bounds.Bottom;
                largestWidth = Math.Max(largestWidth, addedPara.Bounds.Width);
            }

            foreach (var paragraph in body.Paragraphs)
            {
                SetHorizontalAlignmentPosition();
            }

            if (Paragraphs != null && Paragraphs.Count() > 0)
            {
                Bounds.Height = currentHeight;
            }

            Bounds.Top = GetAlignmentVertical();
        }

        private double GetParagraphAscendantSpacingInPixels(eDrawingTextLineSpacing lineSpacingType, double spacingValue, ITextShaper fmExact, float fontSize, out double multiplier)
        {
            if (lineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                multiplier = -1;
                return spacingValue.PointToPixel();
            }
            else
            {
                multiplier = (spacingValue / 100);
                return multiplier * fmExact.GetAscentInPoints(fontSize).PointToPixel();
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
        //public void SetMeasurer(FontMeasurerTrueType fontMeasurer)
        //{
        //    _measurer = fontMeasurer;
        //}

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
                    alignmentY = Bounds.Top;
                    break;
                    //Center means center of a Shape's ENTIRE bounding box height.
                    //Not center of the Inset GetRectangle
                case eTextAnchoringType.Center:
                    alignmentY = (MaxHeight - Bounds.Height)/2 + Bounds.Top;
                    break;
                case eTextAnchoringType.Bottom:
                    alignmentY = MaxHeight - Bounds.Height;
                    break;
            }

            _alignmentY = alignmentY;

            return _alignmentY.Value;
        }


        /// <summary>
        /// Each file format defines its own paragraph
        /// </summary>
        /// <returns></returns>
        internal abstract ParagraphItem CreateParagraph(TextBodyItem textBody,  ExcelDrawingParagraph paragraph, BoundingBox parent, string textIfEmpty="");

        internal abstract ParagraphItem CreateParagraph(TextBodyItem textBody, BoundingBox parent, string textIfEmpty = "");

        /// <summary>
        /// Each file format defines its own paragraph
        /// </summary>
        /// <returns></returns>
        internal abstract ParagraphItem CreateParagraph(TextBodyItem textBody, BoundingBox parent);
    }
}
