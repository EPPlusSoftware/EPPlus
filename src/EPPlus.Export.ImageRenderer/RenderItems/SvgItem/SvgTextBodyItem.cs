using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBodyItem : TextBody
    {
        public SvgTextBodyItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {
        }

        internal override List<ParagraphContainer> Paragraphs { get; set; } = new List<ParagraphContainer>();

        public override void Render(StringBuilder sb)
        {
            sb.AppendLine($"<g transform=\"translate({Bounds.GlobalX.ToString(CultureInfo.InvariantCulture)},{Bounds.GlobalY.ToString(CultureInfo.InvariantCulture)})\" ");
            //base.Render(sb);
            sb.Append(" >");
            sb.AppendLine($"<title>txtBody</title>");

            //var bb = new SvgRenderRectItem(Bounds);
            //bb.X = 0;
            //bb.Width = Width;
            //bb.Height = Height;
            //bb.FillColor = "red";
            //bb.FillOpacity = 0.5;
            //bb.Render(sb);

            foreach (var item in Paragraphs)
            {
                item.Render(sb);
            }
            sb.AppendLine("</g>");
        }

        internal override ParagraphContainer CreateParagraph(BoundingBox parent)
        {
            return new SvgParagraphItem(DrawingRenderer, parent);
        }

        internal override ParagraphContainer CreateParagraph(ExcelDrawingParagraph paragraph, BoundingBox parent)
        {
            return new SvgParagraphItem(DrawingRenderer, parent, paragraph);
        }
    }
}
