using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgParagraphRenderer : SvgBaseRenderer<ParagraphRenderItem> 
    {
        public SvgParagraphRenderer(StringBuilder outputStream) : base(outputStream)
        {

        }
        internal string Suffix = "px";
        public override void Render(ParagraphRenderItem item)
        {
            var sb = OutputStream;
            var fontSize = item.ParagraphFont.Size.PointToPixel().ToString(CultureInfo.InvariantCulture);

            sb.AppendLine($"<g transform=\"translate({item.Bounds.Left.PointToPixelString()},{item.Bounds.Top.PointToPixelString()})\" >");

            sb.AppendLine("<title>paragraph</title> ");

            //if (DisplayBounds)
            //{
            //    sb.AppendLine($"<g>");
            //    sb.AppendLine("<title>Bounding-Box: Paragraph</title> ");
            //    SvgRenderRectItem visualBoundingBox = new SvgRenderRectItem(DrawingRenderer, ParentTextBody.Bounds);

            //    //Left/Top handled by transform
            //    visualBoundingBox.Bounds.Width = item.Bounds.Width;
            //    visualBoundingBox.Bounds.Height = itemBounds.Height;

            //    visualBoundingBox.FillOpacity = 0.3;
            //    visualBoundingBox.FillColor = "red";
            //    visualBoundingBox.Render(sb);
            //    sb.AppendLine($"</g>");

            //}

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
            RenderBase(item);

            sb.Append(/*$"{GetHorizontalAlignmentAttribute(Bounds.X)} y=\"{Bounds.Y}\" " +*/
                $"font-family=\"{item.ParagraphFont.FontFamily},{item.ParagraphFont.FontFamily}_MSFontService,sans-serif\" " +
                $"font-size=\"{fontSize.ToString(CultureInfo.InvariantCulture)}px\" >");

            if (item.Runs != null && item.Runs.Count > 0)
            {
                foreach (var textRun in item.Runs)
                {
                    RenderTextRun(textRun);
                }
            }

            sb.AppendLine("</text>");
            sb.AppendLine("</g>");

            RenderBase(item);
            OutputStream.AppendFormat("/>");
        }

        private void RenderTextRun(TextRunRenderItem textRun)
        {
            throw new NotImplementedException();
        }
    }
}
