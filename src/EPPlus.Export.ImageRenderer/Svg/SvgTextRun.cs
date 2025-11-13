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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgTextRun : SvgRenderItem
    {
        double LineSpacingPerNewLine;
        double _yPosition;
        double ClippingHeight;
        double fontSizeInPixels;
        eTextAlignment horizontalTextAlignment;

        internal RectBase BoundingBox = new RectBase();
        internal Coordinate origin = new Coordinate(0, 0);
        List<string> Lines;

        MeasurementFont measurementFont;

        FontMeasurerTrueType fmExact;

        //Unnecesary??
        double TextLengthInPixels;
        double BaselineSpacing;

        string horizontalAttribute;
        string originalText;

        double _xPosition;

        string fontStyleAttributes;

        /// <summary>
        /// constructor. enter linespacing in pixels
        /// </summary>
        /// <param name="textRun"></param>
        /// <param name="lineSpacing"></param>
        internal SvgTextRun(ExcelParagraphTextRunBase textRun, double lineSpacing, double textMaxX, double textMaxY, double xPosition, double yPosition, double baselineLineSpacing = double.NaN) : base()
        {
            originalText = textRun.Text;
            Lines = SplitIntoLines(originalText);

            measurementFont = textRun.GetMeasureFont();

            if(textRun.HasFill())
            {
                FillColor = textRun.Fill.Color.To6CharHexString();
            }

            fmExact = new FontMeasurerTrueType(measurementFont);

            horizontalTextAlignment = textRun.Paragraph.HorizontalAlignment;

            LineSpacingPerNewLine = lineSpacing;

            //Used for the first line
            BaselineSpacing = baselineLineSpacing;

            _xPosition = xPosition;
            _yPosition = yPosition;

            //origin.X = xPosition;
            //origin.Y = yPosition;

            fontSizeInPixels = ((double)measurementFont.Size).PointToPixel(true);

            ClippingHeight = textMaxY;
            //No idea how to get this property now since we have no access to textbody
            bool wrapText = true;
            if (wrapText)
            {
                CalculateTextWrapping(textMaxX);
            }

            if (textRun.FontItalic)
            {
                fontStyleAttributes += " font-style=\"italic\" ";
            }
            if(textRun.FontBold)
            {
                fontStyleAttributes += "font-weight=\"bold\" ";
            }
            if(textRun.FontUnderLine != eUnderLineType.None | textRun.FontStrike != eStrikeType.No)
            {

                fontStyleAttributes += "text-decoration=\"";
                if (textRun.FontUnderLine != eUnderLineType.None)
                {
                    switch (textRun.FontUnderLine)
                    {
                        case eUnderLineType.Single:
                            fontStyleAttributes += "underline";
                            break;
                        //These are all css only apparently
                        //case eUnderLineType.Double:
                        //    fontStyleAttributes += "double";
                        //    break;
                        //case eUnderLineType.Dotted:
                        //    fontStyleAttributes += "dotted";
                        //    break;
                        //case eUnderLineType.Dash:
                        //    fontStyleAttributes += "dashed";
                        //    break;
                        //case eUnderLineType.Wavy:
                        //    fontStyleAttributes += "wavy";
                        //    break;
                        default:
                            fontStyleAttributes += "underline";
                            break;
                            //throw new NotImplementedException("Not implemented yet");
                    }
                }
                
                if(textRun.FontStrike == eStrikeType.Single)
                {
                    if(textRun.FontUnderLine != eUnderLineType.None)
                    {
                        fontStyleAttributes += ",";
                    }
                    fontStyleAttributes += "line-through";
                }

                fontStyleAttributes += "\" ";
            }
        }

        internal SvgTextRun(ExcelRichText textRun, double lineSpacing, double textMaxX, double textMaxY, double xPosition, double yPosition, MeasurementFont mf, ExcelHorizontalAlignment horAlign, double baselineLineSpacing = double.NaN) : base()
        {
            originalText = textRun.Text;
            Lines = SplitIntoLines(originalText);

            fmExact = new FontMeasurerTrueType(mf);
            measurementFont = mf;

            horizontalTextAlignment = (eTextAlignment)horAlign;

            LineSpacingPerNewLine = fmExact.GetSingleLineSpacing().PointToPixel(true);

            //Used for the first line
            BaselineSpacing = fmExact.GetBaseLine().PointToPixel(true);

            _xPosition = xPosition;
            _yPosition = yPosition;

            fontSizeInPixels = ((double)mf.Size).PointToPixel(true);

            ClippingHeight = textMaxY;
            //No idea how to get this property now since we have no access to textbody
            bool wrapText = true;
            if (wrapText)
            {
                CalculateTextWrapping(textMaxX);
            }
        }

        internal void AdjustLineSpacing(double lineMultiplier)
        {
            LineSpacingPerNewLine = lineMultiplier * fmExact.GetSingleLineSpacing().PointToPixel(true);
        }

        public RectBase textArea;

        public override SvgItemType Type => SvgItemType.TSpan;

        public override void Render(StringBuilder sb)
        {
            string finalString = "";
            bool isFirst = true;
            bool useBaselineSpacing = double.IsNaN(BaselineSpacing) == false;

            foreach (var line in Lines)
            {
                var yIncrease = isFirst && useBaselineSpacing ? BaselineSpacing : LineSpacingPerNewLine;
                isFirst = false;

                yIncrease = EPPlus.Fonts.OpenType.Utils.TextUtils.RoundToWhole(yIncrease);

                _yPosition += yIncrease;
                string visibility = "";
                if (_yPosition >= ClippingHeight)
                {
                    visibility = "display=\"none\"";
                }

                var yIncreaseString = yIncrease.ToString(CultureInfo.InvariantCulture);

                finalString += $"<tspan {visibility} x=\"{(_xPosition).ToString(CultureInfo.InvariantCulture)}\" dy=\"{yIncreaseString}px\" " + $"{fontStyleAttributes}" ;
                if (measurementFont != null)
                {
                    finalString += $"font-family=\"{measurementFont.FontFamily},{measurementFont.FontFamily}_MSFontService,sans-serif\" " + $"font-size=\"{fontSizeInPixels.ToString(CultureInfo.InvariantCulture)}px\" ";
                }
                sb.Append(finalString);
                //Get color etc.
                base.Render(sb);
                finalString = "";

                finalString += ">";
                finalString += line;
                finalString += "</tspan>";
            }

            sb.Append(finalString);
            //throw new NotImplementedException();
        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }

        private int GetNumberOfLines()
        {
            if (originalText != null)
            {
                SplitIntoLines(originalText);
                return Lines.Count;
            }
            else
            {
                return 0;
            }
        }

        internal int GetLineCount()
        {
            if (Lines == null) return 0;

            return Lines.Count;
        }

        internal List<string> SplitIntoLines(string text)
        {
            return text.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList();
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = origin.X;
            it = origin.Y;

            ir = CalculateRightPositionInPixels();
            ib = CalculateBottomPositionInPixels();

            BoundingBox.Left = il;
            BoundingBox.Top = it;
            BoundingBox.Right = ir;
            BoundingBox.Bottom = ib;
        }

        internal double CalculateRightPositionInPixels()
        {
            var numLines = GetNumberOfLines();
            double retPos = origin.X;

            if (numLines <= 0 || string.IsNullOrEmpty(originalText))
            {
                TextLengthInPixels = 0;
                return retPos;
            }

            double longestWidth = -1;

            for (int i = 0; i < numLines; i++)
            {
                var text = Lines[i];
                var width = CalculateTextWidth(Lines[i]);

                if (width > longestWidth)
                {
                    longestWidth = width;
                }
            }

            TextLengthInPixels = longestWidth;

            retPos += longestWidth;

            return retPos;
        }
        internal double CalculateBottomPositionInPixels()
        {
            var lineSize = ((double)measurementFont.Size).PointToPixel(true) + origin.Y;
            var bottomPosition = lineSize * GetNumberOfLines();
            return bottomPosition;
        }

        internal double CalculateTextWidth(string targetString)
        {
            var textMesurer = new FontMeasurerTrueType(measurementFont);
            textMesurer.MeasureWrappedTextCells = true;
            var width = textMesurer.MeasureTextWidth(targetString);

            return width;
        }

        internal void CalculateTextWrapping(double maxWidth)
        {
            List<string> NewContentLines = new List<string>();
            var textMesurer = new FontMeasurerTrueType(measurementFont);
            var newLines = textMesurer.MeasureAndWrapText(originalText, measurementFont, maxWidth);
            Lines = newLines;
        }
    }
}