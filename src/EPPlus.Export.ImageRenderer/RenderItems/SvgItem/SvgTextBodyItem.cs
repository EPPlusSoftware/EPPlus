using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBodyItem : TextBody<SvgParagraphItem>
    {
        public SvgTextBodyItem(BoundingBox parent) : base(parent)
        {
        }

        internal override List<SvgParagraphItem> Paragraphs { get; set; } = new List<SvgParagraphItem>();

        public override void Render(StringBuilder sb)
        {
            sb.AppendLine($"<g transform=\"translate({Bounds.GlobalX},{Bounds.GlobalY})\" >");
            sb.AppendLine($"<title>");
            sb.AppendLine($"txtBody");
            sb.AppendLine("</title>");

            foreach(var item in Paragraphs)
            {
                item.Render(sb);
            }
            sb.AppendLine("</g>");

            //var textBodyGroup = new SvgElement("g");
            //textBodyGroup.AddAttribute("transform", $"translate({Bounds.X},{Bounds.Y})");

            //var txtBodyTitle = new SvgElement("title");
            //txtBodyTitle.Content = "txtBody";
            //textBodyGroup.AddChildElement(txtBodyTitle);


            //var txBodyVisual = new SvgElement("use");
            //txBodyVisual.AddAttribute("href", "#defaultRect");
            //txBodyVisual.AddAttribute("fill", "green");
            //txBodyVisual.AddAttribute("opacity", "0.5");

            //textBodyGroup.AddChildElement(txBodyVisual);

            //var str = RenderSvgElement(textBodyGroup);

            //sb.AppendLine(str);

            //foreach (var item in Paragraphs)
            //{
            //    item.Render(sb);
            //}

            //sb.AppendLine("</g>");

            //foreach(var paragraph in Paragraphs)
            //{

            //}
            //shapeRoot.AddChildElement(textBodyGroup);
        }
    }
}
