using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgParagraphItem : ParagraphContainer
    {
        public SvgParagraphItem()
        {
        }

        public SvgParagraphItem(BoundingBox parent) : base(parent)
        {
        }

        public SvgParagraphItem(ExcelDrawingParagraph p, BoundingBox parent) : base(p, parent)
        {
            //---Add Actual textruns---
            for (int i = 0; i < p.TextRuns.Count; i++)
            {
                AddTextRun(p.TextRuns[i], _textRunContent[i]);
            }
        }

        internal protected override void AddTextRun(ExcelParagraphTextRunBase origTxtRun, string runDisplayString)
        {
            AddTextRun<SvgTextRunItem>(origTxtRun, runDisplayString);
        }

        private string GetHorizontalAlignmentAttribute(double indentX)
        {
            string ret = "";
            var xStr = indentX.ToString(CultureInfo.InvariantCulture);

            switch (_hAlign)
            {
                default:
                case eTextAlignment.Left:
                    ret = $"text-anchor=\"start\" x=\"{xStr}\" ";
                    break;
                case eTextAlignment.Center:
                    ret = $"text-anchor=\"middle\" x=\"{xStr}\" ";
                    break;
                case eTextAlignment.Right:

                    ret = $"text-anchor=\"end\" x=\"{xStr}\" ";
                    break;
            }

            return ret;
        }

        //private string GetVerticalAlignAttribute(double textOrigY)
        //{
        //    string ret = "dominant-baseline=";

        //    switch (VerticalAlignment)
        //    {
        //        case eTextAnchoringType.Top:
        //            ret += $"\"text-top\" ";
        //            break;
        //        case eTextAnchoringType.Center:
        //            ret += $"\"middle\" ";
        //            break;
        //        case eTextAnchoringType.Bottom:
        //            ret += $"\"text-bottom\" ";
        //            break;
        //        default:
        //            return "";
        //    }

        //    var yStrValue = textOrigY.ToString(CultureInfo.InvariantCulture);
        //    ret += $" y=\"{yStrValue}\"";

        //    return ret;
        //}

        public override void Render(StringBuilder sb)
        {
            var fontSize = _paragraphFont.Size.PointToPixel().ToString(CultureInfo.InvariantCulture);

            sb.AppendLine($"<g transform=\"translate({Bounds.X},{Bounds.Y})\" >");

            sb.AppendLine("<title>");
            sb.AppendLine("paragraph");
            sb.AppendLine("</title>");


            sb.Append("<text ");
            base.Render(sb);

            sb.Append(/*$"{GetHorizontalAlignmentAttribute(Bounds.X)} y=\"{Bounds.Y}\" " +*/
                $"font-family=\"{_paragraphFont.FontFamily},{_paragraphFont.FontFamily}_MSFontService,sans-serif\" " +
                $"font-size=\"{fontSize}px\" >");

            if (_textRunItems.Count > 0)
            {
                foreach (var textRun in _textRunItems)
                {
                    textRun.Render(sb);
                }
            }

            sb.AppendLine("</text>");
            sb.AppendLine("</g>");
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new System.NotImplementedException();
        }
    }
}
