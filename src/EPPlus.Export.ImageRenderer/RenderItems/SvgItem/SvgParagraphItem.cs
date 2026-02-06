using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgParagraphItem : ParagraphItem
    {
        public override RenderItemType Type => RenderItemType.Paragraph;

        public SvgParagraphItem(TextBodyItem textBody, DrawingBase renderer, BoundingBox parent) : base(textBody, renderer, parent)
        {
        }

        public SvgParagraphItem(TextBodyItem textBody, DrawingBase renderer, BoundingBox parent, ExcelDrawingParagraph p, string textIfEmpty = null) : base(textBody, renderer, parent, p, textIfEmpty)
        {
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

            sb.AppendLine($"<g transform=\"translate({Bounds.Left.PointToPixelString()},{Bounds.Top.PointToPixelString()})\" >");

            sb.AppendLine("<title>paragraph</title> ");

            //var bb = new SvgRenderRectItem(DrawingRenderer, Bounds);
            ////The bb is affected by the Transform so set pos to zero
            //if (IsFirstParagraph == false)
            //{
            //    bb.Y = 0;
            //}
            //bb.X = 0;

            //bb.Width = Bounds.Width;
            //bb.Height = Bounds.Height;
            //bb.FillColor = FillColor;
            //bb.FillOpacity = 0.3;
            //bb.Render(sb);

            //Render text run debug boxes as we cannot place the rects after or inside the text element

            //if (Runs.Count > 0)
            //{
            //    //double lastWidth = 0;

            //    var bbLines = new SvgRenderRectItem();

            //    foreach (var textRun in Runs)
            //    {
            //        ////render txtRun debug
            //        bbLines.X = textRun.Bounds.Left;
            //        bbLines.Width = textRun.Bounds.Width;
            //        if (IsFirstParagraph == false)
            //        {
            //            bbLines.Y = -textRun.FontSizeInPixels.PixelToPoint();
            //        }

            //        //Render each line debug
            //        bbLines.FillColor = "purple";
            //        for (int i = 0; i < textRun.Lines.Count; i++)
            //        {
            //            if (i > 0)
            //            {
            //                bbLines.Y += textRun.YIncreasePerLine[i];
            //                //lastWidth = 0;
            //                bbLines.X = 0;
            //            }
            //            else
            //            {
            //                bbLines.X = textRun.Bounds.Left;
            //            }

            //            bbLines.Height = textRun.FontSizeInPixels;
            //            bbLines.Width = textRun.PerLineWidth[i];
            //            bbLines.RenderRect(sb);
            //        }
            //    }
            //}

            sb.Append("<text ");
            SvgBaseRenderer.BaseRender(sb, this);
            //Render(sb);

            sb.Append(/*$"{GetHorizontalAlignmentAttribute(Bounds.X)} y=\"{Bounds.Y}\" " +*/
                $"font-family=\"{_paragraphFont.FontFamily},{_paragraphFont.FontFamily}_MSFontService,sans-serif\" " +
                $"font-size=\"{fontSize}px\" >");

            if (Runs != null && Runs.Count > 0)
            {
                foreach (var textRun in Runs)
                {
                    textRun.Render(sb);
                }
            }

            sb.AppendLine("</text>");
            sb.AppendLine("</g>");
        }
        internal override TextRunItem CreateTextRun(ExcelParagraphTextRunBase run, BoundingBox parent, string displayText)
        {
            return new SvgTextRunItem(DrawingRenderer, parent, run, displayText);
        }
        internal override TextRunItem CreateTextRun(string text, ExcelTextFont font, BoundingBox parent, string displayText)
        {
            return new SvgTextRunItem(DrawingRenderer, parent, text, font, displayText);
        }
    }
}
