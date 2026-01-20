using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
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

            sb.AppendLine("<title>paragraph</title> ");

            var bb = new SvgRenderRectItem();
            //The bb is affected by the transform so set pos to zero
            if(IsFirstParagraph == false)
            {
                bb.Y = -_paragraphFont.Size.PointToPixel();
            }
            bb.X = 0;

            bb.Width = Bounds.Width;
            bb.Height = Bounds.Height;
            bb.FillColor = "gold";
            bb.FillOpacity = 0.3;
            bb.Render(sb);

            //Render text run debug boxes as we cannot place the rects after or inside the text element

            if (_textRunItems.Count > 0)
            {
                //double lastWidth = 0;

                var bbLines = new SvgRenderRectItem();

                foreach (var textRun in _textRunItems)
                {
                    ////render txtRun debug
                    bbLines.X = textRun.Bounds.Left;
                    bbLines.Width = textRun.Bounds.Width;
                    if (IsFirstParagraph == false)
                    {
                        bbLines.Y = -textRun.FontSizeInPixels.PixelToPoint();
                    }

                    //Render each line debug
                    bbLines.FillColor = "purple";
                    for (int i = 0; i < textRun.Lines.Count; i++)
                    {
                        if (i > 0)
                        {
                            bbLines.Y += textRun.YIncreasePerLine[i];
                            //lastWidth = 0;
                            bbLines.X = 0;
                        }
                        else
                        {
                            bbLines.X = textRun.Bounds.Left;
                        }

                        bbLines.Height = textRun.FontSizeInPixels;
                        bbLines.Width = textRun.PerLineWidth[i];
                        bbLines.RenderRect(sb);
                    }
                }
            }

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

        internal override TextRunItem CreateTextRun(ExcelParagraphTextRunBase run, BoundingBox parent, string displayText)
        {
            return new SvgTextRunItem(run, parent, displayText);
        }
    }
}
