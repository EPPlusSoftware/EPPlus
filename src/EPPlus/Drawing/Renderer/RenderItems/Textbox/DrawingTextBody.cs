using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing.Style.Font;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using System;
using System.Globalization;
using System.Linq;
using EPPlus.Fonts.OpenType.TextShaping;
namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    public class DrawingTextBody : RenderTextBody
    {
        internal ExcelDrawing _drawing;
        internal MeasurementFont _defaultFont;
        internal ExcelTheme Theme { get; private set; }

        public DrawingTextBody(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, MeasurementFont defaultFont, bool autoSize, bool clampedToParent = false) : base(renderContext, parent, autoSize)
        {
            Init(drawing, defaultFont);
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
        }
        public DrawingTextBody(RenderContext renderContext, ExcelDrawing drawing, BoundingBox parent, MeasurementFont defaultFont, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize=false) : base(renderContext, parent, autoSize)
        {
            Init(drawing, defaultFont);
            Left = left;
            Top = top;
            Width = maxWidth;
            Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
        }

        private void Init(ExcelDrawing drawing, MeasurementFont defaultFont)
        {
            _drawing = drawing;
            Theme = drawing._drawings.Worksheet.Workbook.ThemeManager.GetOrCreateTheme();
            _defaultFont = defaultFont;
            if (defaultFont == null || string.IsNullOrEmpty(_defaultFont.FontFamily))
            {
                _defaultFont.FontFamily = Theme.FontScheme.MinorFont.FirstOrDefault(x => x is ExcelDrawingFontSpecial fs && fs.Type == eFontType.Latin)?.Typeface ?? "Aptos Narrow";
                _defaultFont.FontFamily = RenderContext.GetFamilyForTarget(_defaultFont.FontFamily);
            }
        }

        public void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null, bool applyDefaultMargins = false)
        {
            bool isFirst = Paragraphs.Count == 0;
            Text = text;
            var paragraph = CreateParagraph(this, item, this, text);
            paragraph.Name = $"Paragraph{Paragraphs.Count}";
            paragraph.Top = startingY;

            if (AutoSize)
            {
                if (Paragraphs.Count == 0)
                {
                    Height = paragraph.Height;
                    if(Height == 0)
                    {
                        Height = 3d;
                    }
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
        public void AddParagraph(string text, ExcelDrawingParagraph paragraph)        
        { 
            if(paragraph==null)
            {
                AddParagraph(text);
            }
            else
            {
                ImportParagraph(paragraph, 0, text);
            }
        }
        //Horizontal alignment should technically be set directly in paragraph
        //But as the text is not measured until a paragraph has been imported an autosized textbody
        //does not know its maximum size until all its paragraphs has been imported
        //thus after all have we need to adjust.
        //The alternative would be to perform the performance heavy measurement twice.
        internal void SetHorizontalAlignmentPosition()
        {
            foreach (var p in Paragraphs)
            {
                switch (p.HorizontalAlignment)
                {
                    case TextAlignment.Left:
                        p.Left = 0;
                        break;
                    case TextAlignment.Center:
                        p.Left = (Width / 2) - (p.Width / 2);
                        break;
                    case TextAlignment.Right:
                        p.Left = Right - p.Width;
                        break;
                    case TextAlignment.Distributed:
                    case TextAlignment.Justified:
                    case TextAlignment.JustifiedLow:
                    case TextAlignment.ThaiDistributed:
                        p.Left = 0;                    //TODO: Set left for now as we do not support distributed spacing yet
                        break;
                }
            }
        }

        internal TextAlignment TranslateHorizontalPosition(ExcelHorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case ExcelHorizontalAlignment.Left:
                    return TextAlignment.Left;
                case ExcelHorizontalAlignment.Center:
                    return TextAlignment.Center;
                case ExcelHorizontalAlignment.Right:
                    return TextAlignment.Right;
                case ExcelHorizontalAlignment.Distributed:
                case ExcelHorizontalAlignment.CenterContinuous:
                case ExcelHorizontalAlignment.Justify:
                case ExcelHorizontalAlignment.General:
                default:
                    return TextAlignment.Left;          //TODO: Set left for now as we do not support distributed spacing yet
            }
        }

        internal virtual void ImportTextBodyAndParagraphs(ExcelTextBody body, ExcelHorizontalAlignment horizontalDefault = ExcelHorizontalAlignment.Left)
        {
            Text = null;
            VerticalAlignment = (TextAnchoringType)body.Anchor;

            //We already apply bounds top via the parent Transform
            double currentHeight = 0;
            double largestWidth = double.MinValue;

            body.GetInsetsInPoints(out double left, out double top, out double right, out double bottom);

            if (AutoSize == false)
            {
                LeftMargin = left;
                TopMargin = top;
                RightMargin = right;
                BottomMargin = bottom;

                MaxHeight = MaxHeight - top - bottom;
                MaxWidth = MaxWidth - left - right;
                Height = MaxHeight;
                Width = MaxWidth;
            }

            foreach (var paragraph in body.Paragraphs)
            {

                ImportParagraph(paragraph, currentHeight);
                var addedPara = Paragraphs.Last();

                currentHeight = addedPara.Bottom;
                largestWidth = Math.Max(largestWidth, addedPara.Width);
            }


            if (Paragraphs != null && Paragraphs.Count() > 0 && AutoSize)
            {
                Height = currentHeight;
            }

            //Ensure contentBounds are calculated and paragraphs don't overlap
            RecalculateParagraphs();

            //Alignment adjustment for e.g. ChartTitles one paragraph may be longer than another
            //Therefore as paragraphs have no awareness of eachother we must compare and adjust
            foreach (var paragraph in body.Paragraphs)
            {
                SetHorizontalAlignmentPosition();
            }

            Top = GetAlignmentVertical();

            //Style.FillColor = FontColorString;
        }

        internal DrawingParagraphRenderItem CreateParagraph(DrawingTextBody textBody, BoundingBox parent)
        {
            return new DrawingParagraphRenderItem(RenderContext, textBody, parent);
        }

        internal DrawingParagraphRenderItem CreateParagraph(DrawingTextBody textBody, ExcelDrawingParagraph paragraph, BoundingBox parent, string textIfEmpty = null)
        {
            return new DrawingParagraphRenderItem(RenderContext, textBody, parent, paragraph, textIfEmpty);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="textIfEmpty"></param>
        /// <returns></returns>
        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, string textIfEmpty = "")
        {
            return new DrawingParagraphRenderItem(RenderContext, this, parent, textIfEmpty);
        }

        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, IRichTextFormatSimple richText)
        {
            var paragraph = new SvgParagraphRenderItem(RenderContext, this, parent, "");
            paragraph.AddRichText(richText);
            return paragraph;
        }
        internal ParagraphRenderItem AddParagraph(double value, ExcelFormatTranslator nf, double maxWidth,
                                                  ExcelDrawingParagraph style = null)
        {
            var mf = style != null && style.DefaultRunProperties.Size >= 0
                   ? style.DefaultRunProperties.GetMeasureFont(_defaultFont) : _defaultFont;
            var font = new FontFormatBase(mf);
            font.Family = RenderContext.GetFamilyForTarget(font.Family);
            var shaper = (ITextShaper)RenderContext.FontEngine.GetShaperForFont(font);

            string chosen = null;
            foreach (var s in NumberFitFormatter.GetCandidates(value, nf))
            {
                chosen = s;
                if (shaper.MeasureTextInPoints(s, font.Size) <= maxWidth) break;
            }
            if (style == null) AddParagraph(chosen); else ImportParagraph(style, 0, chosen);
            return Paragraphs[Paragraphs.Count - 1];
        }
    }

}
