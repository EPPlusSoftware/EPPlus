/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgParagraph : SvgRenderItem
    {
        /// <summary>
        /// -1 if wrapping is never applied
        /// </summary>
        int numLines = -1;

        public override void Render(StringBuilder sb)
        {
            sb.Append("<text ");
            base.Render(sb);

            sb.Append($" {vertAlignAttribute} " + GetHorizontalAlignmentAttribute(XPos) +
                $"font-family=\"{_measurementFont.FontFamily},{_measurementFont.FontFamily}_MSFontService,sans-serif\" " +
                $"font-size=\"{_measurementFont.Size.PointToPixel().ToString(CultureInfo.InvariantCulture)}px\">");

            if (TextRuns.Count > 0)
            {
                foreach (var textRun in TextRuns)
                {
                    textRun.Render(sb);
                }
            }

            sb.Append("</text>");
        }

        public override SvgItemType Type => SvgItemType.Text;

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = ParagraphArea.Left + LeftMargin;
            it = ParagraphArea.Top;
            ir = ParagraphArea.Right - RightMargin;
            ib = ParagraphArea.Bottom;
        }

        string vertAlignAttribute = "";
        protected List<SvgTextRun> TextRuns = new List<SvgTextRun>();

        //ExcelDrawingParagraph Paragraph;
        double RightMargin;
        double LeftMargin;

        eTextAlignment HorizontalAlignment;

        RectBase ParagraphArea;

        /// <summary>
        /// X position
        /// </summary>
        protected double XPos;

        /// <summary>
        /// Line-Spacing in pixels
        /// </summary>
        protected double LineSpacing;
        protected double LineSpacingAscendantOnly;

        MeasurementFont _measurementFont;

        bool IsFirstParagraph = false;

        ITextMeasurerWrap fmtt;

        eDrawingTextLineSpacing lnType;
        double? lnMultiplier = null;
        double paragraphHeight;
        private string text;
        private ExcelTextFont font;
        private RectBase area;
        private double posY;
        /// <summary>
        /// First paragraph must use different linespacing
        /// </summary>
        /// <param name="p">paragraph data. Ideally only used in constructor as datasource</param>
        /// <param name="paragraphArea"></param>
        /// <param name="fmtt"></param>
        /// <param name="VertAlign"></param>
        /// <param name="clippingHeight"></param>
        /// <param name="isFirstParagraph"></param>
        public SvgParagraph(ExcelDrawingParagraph p, RectBase paragraphArea, string VertAlign, double yPosition, bool isFirstParagraph = false)
        {
            fmtt = p._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;

            _measurementFont = p.DefaultRunProperties.GetMeasureFont();
            fmtt.SetFont(_measurementFont);
            vertAlignAttribute = VertAlign;

            //Seperated out in case of some final render item
            //needing to adjust without changing the original values
            RightMargin = p.RightMargin;

            //p.Indent
            //0.5 inches per indent level 48 pixels

            var indent = 48 * p.IndentLevel;
            LeftMargin = p.LeftMargin + p.Indent + indent;

            ParagraphArea = paragraphArea;
            HorizontalAlignment = p.HorizontalAlignment;

            //Must be set before linespacing
            IsFirstParagraph = isFirstParagraph;

            LineSpacing = GetParagraphLineSpacingInPixels(p, fmtt);

            XPos = GetAlignmentHorizontal(HorizontalAlignment);

            GetBounds(out double l, out double t, out double r, out double b);
            var textMaxWidth = r - l;

            //paragraphHeight = p.GetParagraphHeightInPixels(fmtt, textMaxWidth.PixelToPoint());

            lnType = p.LineSpacing.LineSpacingType;

            foreach (var run in p.TextRuns)
            {
                AddTextRun(run, paragraphArea.Bottom, yPosition);
            }

            if (p._paragraphs.WrapText == eTextWrappingType.Square)
            {
                List<string> textFragments = new List<string>();
                List<MeasurementFont> fonts = new List<MeasurementFont>();

                foreach (var txtRun in p.TextRuns)
                {
                    textFragments.Add(txtRun.Text);
                    fonts.Add(txtRun.GetMeasurementFont());
                }

                var trueTypeMeasurer = (FontMeasurerTrueType)fmtt;
                var maxWidthPoints = textMaxWidth.PixelToPoint();
                var svgLines = trueTypeMeasurer.WrapMultipleTextFragments(textFragments, fonts, maxWidthPoints);

                numLines = svgLines.Count;

                //var combinedString = string.Join(Environment.NewLine, svgLines.ToArray());
                var combinedString = string.Join(string.Empty, svgLines.ToArray());

                List<string> txtRunStrings = new List<string>();
                //int currentWrappedLine = 0;
                List<int> txtRunStartIndicies = new List<int>();
                List<int> txtRunEndIndicies = new List<int>();
                int lastIndex = 0;
                for (int i = 0; i < TextRuns.Count; i++)
                {
                    var txtString = TextRuns[i].originalText;
                    txtRunStrings.Add(txtString);
                    var indexOfRun = p.Text.IndexOf(txtString, lastIndex);
                    txtRunStartIndicies.Add(indexOfRun);
                    txtRunEndIndicies.Add(indexOfRun + txtString.Length);
                    lastIndex = indexOfRun;
                }

                List<int> lineIndicies = new List<int>();
                lastIndex = 0;
                
                //Last line should be handled by paragraph handling
                for (int i = 0; i< svgLines.Count(); i++)
                {
                    var txtString = svgLines[i];
                    txtRunStrings.Add(txtString);
                    var startIndex = p.Text.IndexOf(txtString, lastIndex);
                    lastIndex = startIndex + txtString.Length;
                    lineIndicies.Add(startIndex + txtString.Length);
                }

                for (int i = 0; i < lineIndicies.Count; i++)
                {
                    var lnBreakPosition = lineIndicies[i];
                    for (int j = 0; j < txtRunEndIndicies.Count; j++)
                    {
                        var start = txtRunStartIndicies[j];
                        var end = txtRunEndIndicies[j];

                        bool containsBreak = (start <= lnBreakPosition && lnBreakPosition < end);
                        if (containsBreak)
                        {
                            var localLnBreakPosition = lnBreakPosition - start;
                            TextRuns[j].InsertLineBreak(localLnBreakPosition);
                            break;
                        }
                    }
                }
            }
            else
            {
                numLines = p.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).Count();
            }
        }

        

        /// <summary>
        /// First paragraph must use different linespacing
        /// </summary>
        /// <param name="p">paragraph data. Ideally only used in constructor as datasource</param>
        /// <param name="paragraphArea"></param>
        /// <param name="fmtt"></param>
        /// <param name="VertAlign"></param>
        /// <param name="clippingHeight"></param>
        /// <param name="isFirstParagraph"></param>
        public SvgParagraph(string text, ExcelTextFont font, RectBase paragraphArea, string VertAlign, double yPosition)
        {
            fmtt = font.PictureRelationDocument.Package.Settings.TextSettings.GenericTextMeasurerTrueType; 

            _measurementFont = font.GetMeasureFont();
            fmtt.SetFont(_measurementFont);
            vertAlignAttribute = VertAlign;
            
            //p.Indent
            //0.5 inches per indent level 48 pixels

            ParagraphArea = paragraphArea;
            HorizontalAlignment = eTextAlignment.Right;

            LineSpacingAscendantOnly = fmtt.GetBaseLine().PointToPixel();
            LineSpacing = fmtt.GetSingleLineSpacing().PointToPixel();

            //Must be set before linespacing
            IsFirstParagraph = true;

            XPos = GetAlignmentHorizontal(HorizontalAlignment);

            //GetBounds(out double l, out double t, out double r, out double b);
            //var textMaxWidth = r - l;

            var tr = new SvgTextRun(text, font, LineSpacing, paragraphArea.Bottom, XPos, yPosition, LineSpacingAscendantOnly);
            TextRuns.Add(tr);
        }

        private string GetHorizontalAlignmentAttribute(double indentX)
        {
            string ret = "";
            var xStr = indentX.ToString(CultureInfo.InvariantCulture);

            switch (HorizontalAlignment)
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

        private double GetParagraphLineSpacingInPixels(ExcelDrawingParagraph p, ITextMeasurerWrap fmExact)
        {
            if (p.LineSpacing.LineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                if (IsFirstParagraph)
                {
                    LineSpacingAscendantOnly = p.LineSpacing.Value.PointToPixel();
                }
                return p.LineSpacing.Value.PointToPixel();
            }
            else
            {
                var multiplier = (p.LineSpacing.Value / 100);
                lnMultiplier = multiplier;
                if (IsFirstParagraph)
                {  
                    LineSpacingAscendantOnly = multiplier * fmExact.GetBaseLine().PointToPixel();
                }
                return multiplier * fmExact.GetSingleLineSpacing().PointToPixel();
            }
        }

        internal double GetAlignmentHorizontal(eTextAlignment txAlignment)
        {
            var area = ParagraphArea;
            double x = 0;
            switch (txAlignment)
            {
                case eTextAlignment.Left:
                default:
                    x = area.Left + LeftMargin;
                    break;
                case eTextAlignment.Center:
                    x = (area.Right / 2) + LeftMargin - RightMargin;
                    break;
                case eTextAlignment.Right:
                    x = area.Right - RightMargin;
                    break;
            }

            return TextUtils.RoundToWhole(x);
        }

        internal void AddTextRun(ExcelParagraphTextRunBase txtRun, double clippingHeight, double yPosition)
        {
            GetBounds(out double l, out double t, out double r, out double b);
            var textMaxWidth = r - l;

            var lineSpacing = LineSpacing;

            SvgTextRun textRun;


            if (TextRuns.Count == 0 && IsFirstParagraph == true)
            {
                textRun = new SvgTextRun(txtRun, lineSpacing, textMaxWidth, clippingHeight, XPos, yPosition, LineSpacingAscendantOnly);
            }
            else
            {
                textRun = new SvgTextRun(txtRun, lineSpacing, textMaxWidth, clippingHeight, XPos, yPosition);

                //If there are multiple sizes/multiple fonts with multiple sizes
                //if (lnType != eDrawingTextLineSpacing.Exactly && txtRun.FontSize != _measurementFont.Size)
                if(lnMultiplier.HasValue)
                {
                    textRun.AdjustLineSpacing(lnMultiplier.Value);
                }
            }

            TextRuns.Add(textRun);
        }

        internal double GetBottomYPosition()
        {
            double bottomY = 0;
            if (IsFirstParagraph)
            {
                bottomY = LineSpacingAscendantOnly + LineSpacing * (numLines - 1);
            }
            else
            {
                bottomY = LineSpacing * numLines;
            }

            ////heigh

            ////var totalY = ParagraphArea.Top + lineSpacingTotal;

            ////if(TextRuns.Count == 0)
            ////{
            ////    return ParagraphArea.Top + LineSpacing;
            ////}

            //if (TextRuns.Count != 0)
            //{
            //    bottomY = TextRuns.Last().getBottomY();
            //}
            //if (IsFirstParagraph)
            //{
            //    bottomY += LineSpacingAscendantOnly + LineSpacing;
            //}
            //else
            //{
            //    bottomY += LineSpacing;
            //}

            return ParagraphArea.Top + bottomY;
        }

        internal void CalculateTextWrapping(double maxWidth, MeasurementFont mFont, string fullParagraphText)
        {
            List<string> NewContentLines = new List<string>();
            fmtt.SetFont(mFont);
            var textWidth = fmtt.MeasureText(fullParagraphText, mFont);
        }
    }
}