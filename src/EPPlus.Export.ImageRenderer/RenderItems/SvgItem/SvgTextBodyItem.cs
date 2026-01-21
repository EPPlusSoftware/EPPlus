using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBodyItem : TextBody
    {
        public SvgTextBodyItem(BoundingBox parent) : base(parent)
        {
        }

        internal override List<ParagraphContainer> Paragraphs { get; set; } = new List<ParagraphContainer>();

        public override void Render(StringBuilder sb)
        {
            sb.AppendLine($"<g transform=\"translate({Bounds.GlobalX},{Bounds.GlobalY})\" ");
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
            return new SvgParagraphItem(parent);
        }

        internal override ParagraphContainer CreateParagraph(ExcelDrawingParagraph paragraph, BoundingBox parent)
        {
            return new SvgParagraphItem(paragraph, parent);
        }
    }
}
