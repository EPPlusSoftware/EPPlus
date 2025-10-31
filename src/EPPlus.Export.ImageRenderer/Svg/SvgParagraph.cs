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
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgParagraph : SvgRenderItem
    {
        public override void Render(StringBuilder sb)
        {
            sb.Append("<text ");
            base.Render(sb);

            sb.Append($" {vertAlignAttribute} " + GetHorizontalAlignmentAttribute(XPos) +
                $"font-family=\"{_MeasurementFont.FontFamily},{_MeasurementFont.FontFamily}_MSFontService,sans-serif\" " +
                $"font-size=\"{_MeasurementFont.Size.PointToPixel().ToString(CultureInfo.InvariantCulture)}px\">");

            foreach (var textRun in TextRuns)
            {
                textRun.Render(sb);
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

        MeasurementFont _MeasurementFont;

        bool IsFirstParagraph = false;

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
            var fmtt = p._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;

            _MeasurementFont = p.DefaultRunProperties.GetMeasureFont();
            vertAlignAttribute = VertAlign;
            
            //Seperated out in case of some final render item
            //needing to adjust without changing the original values
            RightMargin = p.RightMargin;
           
            LeftMargin = p.LeftMargin + p.Indent;

            ParagraphArea = paragraphArea;
            HorizontalAlignment = p.HorizontalAlignment;

            //Must be set before linespacing
            IsFirstParagraph = isFirstParagraph;

            LineSpacing = GetParagraphLineSpacingInPixels(p, fmtt);

            XPos = GetAlignmentHorizontal(HorizontalAlignment);

            GetBounds(out double l, out double t, out double r, out double b);
            var textMaxWidth = r - l;

            var paragraphHeight = p.GetParagraphHeightInPixels(fmtt, textMaxWidth.PixelToPoint());
            //var paragraphHeight = p.GetParagraphSizeInPixels(textMaxWidth.PixelToPoint());

            foreach (var run in p.TextRuns)
            {
                AddTextRun(run, paragraphArea.Bottom, yPosition);
            }
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
                    x = area.Left;
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
            }

            TextRuns.Add(textRun);
        }

        internal double GetBottomYPosition()
        {
            int numberOfLines = 0;
            foreach (var textRun in TextRuns)
            {
                numberOfLines += textRun.GetLineCount();
            }

            double lineSpacingTotal;
            if(IsFirstParagraph)
            {
                lineSpacingTotal = LineSpacingAscendantOnly + LineSpacing * (numberOfLines - 1);
            }
            else
            {
                lineSpacingTotal = LineSpacing * numberOfLines;
            }

            var totalY = ParagraphArea.Top + lineSpacingTotal;

            return totalY;
        }
    }
}