/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  23/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
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
        IBasicIShapesRenderer<StringBuilder> _shapeRenderer;
        public SvgParagraphRenderer(IBasicIShapesRenderer<StringBuilder> shapeRenderer, StringBuilder outputStream) : base(outputStream)
        {
            _shapeRenderer = shapeRenderer;
        }
        internal SvgFontPlan FontPlan { get; set; } = SvgFontPlan.WithoutContext;
        internal string Suffix = "px";
        public override void Render(ParagraphRenderItem item)
        {
            var sb = OutputStream;
            var fontSize = item.DefaultParagraphFont.Size.PointToPixel().ToString(CultureInfo.InvariantCulture);

            sb.AppendLine($"<g class=\"{item.Name}\" transform=\"translate({item.Left.PointToPixelString()},{item.Top.PointToPixelString()})\" >");

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

            // xml:space="preserve" keeps spaces at the start and end of the runs, which textLength counts.
            // Nothing may be written between the runs: it would be text too.
            sb.Append("<text xml:space=\"preserve\" ");
            RenderBase(item);

            // The default font in the style attribute, as on the runs: an inline style wins over css of a host page.
            sb.Append($"style=\"font-family:{FontPlan.GetFontFamilyAttributeValue(item.DefaultParagraphFont)};font-size:{fontSize}px;\" >");

            if (item.Runs != null && item.Runs.Count > 0)
            {
                foreach (var textRun in item.Runs)
                {
                    _shapeRenderer.Render(textRun);
                }
            }

            sb.AppendLine("</text>");
            sb.AppendLine("</g>");
        }

        //string GetFontStyleAttributes()
        //{
        //    string fontStyleAttributes = " ";

        //    if (_isItalic)
        //    {
        //        fontStyleAttributes += "font-style=\"italic\" ";
        //    }
        //    if (_isBold)
        //    {
        //        fontStyleAttributes += "font-weight=\"bold\" ";
        //    }
        //    if (_underLineType != eUnderLineType.None | _strikeType != eStrikeType.No)
        //    {

        //        fontStyleAttributes += "text-decoration=\" ";
        //        if (_underLineType != eUnderLineType.None)
        //        {
        //            switch (_underLineType)
        //            {
        //                case eUnderLineType.Single:
        //                    fontStyleAttributes += "underline";
        //                    break;
        //                //These are all css only apparently
        //                //case eUnderLineType.Double:
        //                //    fontStyleAttributes += "double";
        //                //    break;
        //                //case eUnderLineType.Dotted:
        //                //    fontStyleAttributes += "dotted";
        //                //    break;
        //                //case eUnderLineType.Dash:
        //                //    fontStyleAttributes += "dashed";
        //                //    break;
        //                //case eUnderLineType.Wavy:
        //                //    fontStyleAttributes += "wavy";
        //                //    break;
        //                default:
        //                    fontStyleAttributes += "underline";
        //                    break;
        //                    //throw new NotImplementedException("Not implemented yet");
        //            }
        //        }

        //        if (_strikeType == eStrikeType.Single)
        //        {
        //            //Has to check if Both underline and strike
        //            if (_underLineType != eUnderLineType.None)
        //            {
        //                fontStyleAttributes += ",";
        //            }
        //            fontStyleAttributes += "line-through";
        //        }

        //        fontStyleAttributes += "\" ";
        //    }

        //    return fontStyleAttributes;
        //}
    }
}
