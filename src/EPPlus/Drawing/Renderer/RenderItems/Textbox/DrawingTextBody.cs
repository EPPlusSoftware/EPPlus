using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
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
    public class DrawingTextbody : EPPlus.Export.ImageRenderer.RenderItems.SvgItem.RenderTextBody
    {
        protected ExcelDrawing _drawing;
        public DrawingTextbody(ExcelDrawing drawing, BoundingBox parent, bool autoSize, bool clampedToParent = false) : base(parent, autoSize)
        {
            _drawing = drawing;
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
        }
        public DrawingTextbody(ExcelDrawing drawing, BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false, bool autoSize=false) : base(parent, autoSize)
        {
            _drawing = drawing;
            Bounds.Left = left;
            Bounds.Top = top;
            Bounds.Width = maxWidth;
            Bounds.Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
        }

        public void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            var measureFont = item.DefaultRunProperties.GetMeasureFont();
            bool isFirst = Paragraphs.Count == 0;

            var paragraph = CreateParagraph(this, item, Bounds, text);
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

        //internal override ParagraphItem CreateParagraph(TextBodyItem textBody, BoundingBox parent)
        //{
        //    return new SvgParagraphItem(this, DrawingRenderer, parent);
        //}

        //internal override ParagraphItem CreateParagraph(TextBodyItem textBody, ExcelDrawingParagraph paragraph, BoundingBox parent, string textIfEmpty = null)
        //{
        //    return new SvgParagraphItem(this, DrawingRenderer, parent, paragraph, textIfEmpty);
        //}

        //internal override ParagraphItem CreateParagraph(TextBodyItem textBody, BoundingBox parent, string textIfEmpty = "")
        //{
        //    return new SvgParagraphItem(this, DrawingRenderer, parent, textIfEmpty);
        //}
    }
}
