using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBodyItem : TextBodyItem
    {
        public SvgTextBodyItem(DrawingBase renderer, BoundingBox parent, bool clampedToParent = false) : base(renderer, parent)
        {
            Bounds.ClampedToParent = clampedToParent;
        }

        internal override List<ParagraphItem> Paragraphs { get; set; } = new List<ParagraphItem>();

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            //sb.AppendLine($"<g transform=\"translate({Bounds.GlobalX.ToString(CultureInfo.InvariantCulture)},{Bounds.GlobalY.ToString(CultureInfo.InvariantCulture)})\" ");
            ////base.Render(sb);
            //sb.Append(" >");
            //sb.AppendLine($"<title>txtBody</title>");
            
            var groupItem = new SvgGroupItem(DrawingRenderer, Bounds, Bounds.Rotation);
            renderItems.Add(groupItem);
            foreach (var item in Paragraphs)
            {
                renderItems.Add(item);
            }
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }

        internal override ParagraphItem CreateParagraph(BoundingBox parent)
        {
            return new SvgParagraphItem(DrawingRenderer, parent);
        }

        internal override ParagraphItem CreateParagraph(ExcelDrawingParagraph paragraph, BoundingBox parent)
        {
            return new SvgParagraphItem(DrawingRenderer, parent, paragraph);
        }
    }
}
