using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    public class DrawingTextBody : RenderTextBody
    {
        internal ExcelDrawing _drawing;

        internal ExcelTheme Theme { get; }

        public DrawingTextBody(ExcelDrawing drawing, BoundingBox parent, bool autoSize, bool clampedToParent = false) : base(parent, autoSize)
        {
            _drawing = drawing;
            Theme = drawing._drawings.Worksheet.Workbook.ThemeManager.GetOrCreateTheme();
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
        }
        public DrawingTextBody(ExcelDrawing drawing, BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize=false) : base(parent, autoSize)
        {
            _drawing = drawing;
            Theme = drawing._drawings.Worksheet.Workbook.ThemeManager.GetOrCreateTheme();
            Bounds.Left = left;
            Bounds.Top = top;
            Bounds.Width = maxWidth;
            Bounds.Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
        }

        public void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            bool isFirst = Paragraphs.Count == 0;
            Text = text;

            var paragraph = CreateParagraph(this, item, Bounds, text);
            paragraph.Bounds.Name = $"Container{Paragraphs.Count}";
            paragraph.Bounds.Top = startingY;

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
            RecalculateParagraphs();
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
                        p.Bounds.Left = 0;
                        break;
                    case TextAlignment.Center:
                        p.Bounds.Left = (Bounds.Width / 2) - (p.Bounds.Width / 2);
                        break;
                    case TextAlignment.Right:
                        p.Bounds.Left = Bounds.Right - p.Bounds.Width;
                        break;
                    case TextAlignment.Distributed:
                    case TextAlignment.Justified:
                    case TextAlignment.JustifiedLow:
                    case TextAlignment.ThaiDistributed:
                        p.Bounds.Left = 0;                    //TODO: Set left for now as we do not support distributed spacing yet
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

            //var defaultAlignment = TranslateHorizontalPosition(horizontalDefault);

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
                //addedPara.HorizontalAlignment = defaultAlignment;

                currentHeight = addedPara.Bounds.Bottom;
                largestWidth = Math.Max(largestWidth, addedPara.Bounds.Width);
            }


            if (Paragraphs != null && Paragraphs.Count() > 0 && AutoSize)
            {
                Bounds.Height = currentHeight;
            }

            //Ensure contentBounds are calculated and paragraphs don't overlap
            RecalculateParagraphs();

            //Alignment adjustment for e.g. ChartTitles one paragraph may be longer than another
            //Therefore as paragraphs have no awareness of eachother we must compare and adjust
            foreach (var paragraph in body.Paragraphs)
            {
                SetHorizontalAlignmentPosition();
            }

            Bounds.Top = GetAlignmentVertical();
        }

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

        internal DrawingParagraphRenderItem CreateParagraph(DrawingTextBody textBody, BoundingBox parent)
        {
            return new DrawingParagraphRenderItem(textBody, parent);
        }

        internal DrawingParagraphRenderItem CreateParagraph(DrawingTextBody textBody, ExcelDrawingParagraph paragraph, BoundingBox parent, string textIfEmpty = null)
        {
            return new DrawingParagraphRenderItem(textBody, parent, paragraph, textIfEmpty);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="textIfEmpty"></param>
        /// <returns></returns>
        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, string textIfEmpty = "")
        {
            return new DrawingParagraphRenderItem(this, parent, textIfEmpty);
        }

        protected override ParagraphRenderItem CreateParagraph(BoundingBox parent, IRichTextFormatSimple richText)
        {
            var paragraph = new SvgParagraphRenderItem(this, parent, "", false);
            paragraph.AddRichText(richText);
            return paragraph;
        }
    }
}
