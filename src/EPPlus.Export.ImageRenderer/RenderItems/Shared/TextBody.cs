using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
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
    internal abstract class TextBody : RenderItem
    {
        //internal BoundingBox Bounds = new BoundingBox();
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

        internal abstract List<ParagraphContainer> Paragraphs { get; set; }

        public bool AllowOverflow;

        internal bool WrapText = true;

        private FontMeasurerTrueType _measurer = null;
        internal string _text;
        public TextBody(DrawingBase renderer, BoundingBox parent)  : base(renderer, parent)
        {
            Bounds.Transform.Name = "TxtBody";

            Bounds.Parent = parent;

            Bounds.Width = parent.Width;
            Bounds.Height = parent.Height;
        }

        public void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text=null)
        {
            var measureFont = item.DefaultRunProperties.GetMeasureFont();
            bool isFirst = Paragraphs.Count == 0;

            var paragraph = CreateParagraph(item, Bounds);
            paragraph.Bounds.Transform.Name = $"Container{Paragraphs.Count}";
            if (string.IsNullOrEmpty(text) == false)
            {
                paragraph.AddText(text, item.DefaultRunProperties);
            }
            paragraph.Bounds.Top = startingY;
            _text = text;
            Paragraphs.Add(paragraph);
        }

        public void ImportTextBody(ExcelTextBody body)
        {
            double l, r, t, b;
            body.GetInsetsOrDefaults(out l, out t, out r, out b);

            LeftMargin = l.PointToPixel();
            TopMargin = t.PointToPixel();
            RightMargin = r.PointToPixel();
            BottomMargin = b.PointToPixel();

            _text = null;
            VerticalAlignment = body.Anchor;
            //We already apply bounds top via the parent Transform
            double paragraphStartY = GetAlignmentVertical();

            foreach (var paragraph in body.Paragraphs)
            {
                if (paragraph == body.Paragraphs[0])
                {
                    //For the first line we always add ascent for Top-aligned but this should not be done for center vertical align
                    //However as paragraph and textRun should not have to deal with vertical align directly
                    //We wish to achieve the effect of not applying dy to the first textrun while not chaning anything
                    //thus: we change paragraphStartY to get around it. Should probably be solved in the textrun somehow 
                    if (VerticalAlignment == eTextAnchoringType.Center && paragraph.TextRuns != null && paragraph.TextRuns.Count > 0)
                    {
                        var _measurer = paragraph._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
                        var spacingValue = GetParagraphAscendantSpacingInPixels(
                            paragraph.LineSpacing.LineSpacingType, paragraph.LineSpacing.Value, _measurer, out double multiplier);

                        if(multiplier == -1)
                        {
                            paragraphStartY -= spacingValue;
                        }
                        else
                        {
                            var runFont = paragraph.TextRuns[0].GetMeasurementFont();
                            var startFont = paragraph.DefaultRunProperties.GetMeasureFont();
                            _measurer.SetFont(runFont);

                            paragraphStartY -= multiplier * _measurer.GetBaseLine().PointToPixel(true);

                            //Reset measurer font
                            _measurer.SetFont(startFont);
                        }


                        //paragraph.TextRuns[0].GetMeasurementFont() *
                        ////var baseLineSize = paragraph.TextRuns[0].FontSize.PointToPixel();
                        ////paragraphStartY = paragraphStartY - baseLineSize;
                    }
                }

                ImportParagraph(paragraph, paragraphStartY);
                var addedPara = Paragraphs.Last();
                paragraphStartY = addedPara.Bounds.Bottom;
            }
            if (Paragraphs != null && Paragraphs.Count() > 0)
            {
                Bounds.Height = paragraphStartY;
            }
        }

        private double GetParagraphAscendantSpacingInPixels(eDrawingTextLineSpacing lineSpacingType, double spacingValue, ITextMeasurerWrap fmExact, out double multiplier)
        {
            if (lineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                multiplier = -1;
                return spacingValue.PointToPixel();
            }
            else
            {
                multiplier = (spacingValue / 100);
                return multiplier * fmExact.GetBaseLine().PointToPixel();
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

        /// <summary>
        /// Same as X or Left
        /// </summary>
        internal double LeftMargin { get { return Bounds.Left; } set { Bounds.Left = value; } }

        /// <summary>
        /// Same as Y or Top
        /// </summary>
        internal double TopMargin { get { return Bounds.Top; } set { Bounds.Top = value; } }

        double _rightMargin = 0;

        /// <summary>
        /// Setting this property also changes the width of the textbody
        /// </summary>
        internal double RightMargin { get { return _rightMargin; } set { _rightMargin = value; Bounds.Width = Bounds.Parent.Width - value; } }

        double _bottomMargin = 0;
        /// <summary>
        /// Setting this property also changes the height of the textbody
        /// </summary>
        internal double BottomMargin { get { return _bottomMargin; } set { _bottomMargin = value; Bounds.Height = Bounds.Parent.Height - value; } }

        public override RenderItemType Type => RenderItemType.Text;

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
                    //Center means center of a Shape's ENTIRE bounding box height.
                    //Not center of the Inset Rectangle
                case eTextAnchoringType.Center:
                    var globalHeight = (DrawingRenderer.Bounds.Height / 2) + Bounds.Top+2;
                    var adjustedHeight = globalHeight - Bounds.GlobalY;

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
        internal abstract ParagraphContainer CreateParagraph(ExcelDrawingParagraph paragraph, BoundingBox parent);

        /// <summary>
        /// Each file format defines its own paragraph
        /// </summary>
        /// <returns></returns>
        internal abstract ParagraphContainer CreateParagraph(BoundingBox parent);
    }
}
