using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBodyItem : TextBodyItem
    {
        public SvgTextBodyItem(DrawingBase renderer, BoundingBox parent, bool autoSize, bool clampedToParent = false) : base(renderer, parent, autoSize)
        {
            //Bounds.ClampedToParent = clampedToParent;
            MaxWidth = parent.Width;
            MaxHeight = parent.Height;
        }
        public SvgTextBodyItem(DrawingBase renderer, BoundingBox parent, double left, double top, double maxWidth, double maxHeight, bool clampedToParent = false) : base(renderer, parent, false)
        {
            Bounds.Left = left;
            Bounds.Top = top;
            Bounds.Width = maxWidth;
            Bounds.Height = maxHeight;
            MaxWidth = maxWidth;
            MaxHeight = maxHeight;
            //Bounds.ClampedToParent = clampedToParent;
        }
        internal override List<ParagraphItem> Paragraphs { get; set; } = new List<ParagraphItem>();

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            SvgGroupItem groupItem;
            if (Bounds.Parent.Rotation == 0) //If the parent is rotated, we should not apply rotation again. This is usually when the parent is a textbox.
            {
                groupItem = new SvgGroupItem(DrawingRenderer, Bounds, Bounds.Rotation);
            }
            else
            {
                groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
            }
            renderItems.Add(groupItem);
            foreach (SvgParagraphItem item in Paragraphs)
            {
                renderItems.Add(item);
            }
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }

        internal override ParagraphItem CreateParagraph(TextBodyItem textBody, BoundingBox parent)
        {
            return new SvgParagraphItem(this, DrawingRenderer, parent);
        }

        internal override ParagraphItem CreateParagraph(TextBodyItem textBody, ExcelDrawingParagraph paragraph, BoundingBox parent, string textIfEmpty = null)
        {
            return new SvgParagraphItem(this, DrawingRenderer, parent, paragraph, textIfEmpty);
        }

        internal override ParagraphItem CreateParagraph(TextBodyItem textBody, BoundingBox parent, string textIfEmpty = "")
        {
            return new SvgParagraphItem(this, DrawingRenderer, parent, textIfEmpty);
        }
    }
}
