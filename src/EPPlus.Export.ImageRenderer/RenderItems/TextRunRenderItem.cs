using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.RenderItems
{
    internal abstract class TextRunRenderItem : RenderItem
    {
        /// <summary>
        /// BoundingBox
        /// </summary>
        internal BoundingBox boundingBox;
        public override RenderItemType Type => RenderItemType.Text;

        internal readonly string originalText;
        private string currentText;

        MeasurementFont measurementFont;
        bool isFirstInParagraph;
        FontMeasurerTrueType Measurer;
        eTextAlignment horizontalTextAlignment;
        MeasurementFontStyles fontStyles;
        double fontSizeInPixels;

        List<string> Lines;

        double _yEndPos;
        double BaselineSpacing;

        bool WrapText;
        double MaxWidthPixels = double.NaN;

        internal TextRunRenderItem(ExcelParagraphTextRunBase run, BoundingBox parent = null)
        {
            WrapText = run.Paragraph._paragraphs.WrapText != eTextWrappingType.None;

            originalText = run.Text;
            currentText = originalText;

            boundingBox = new BoundingBox();
            if (parent != null)
            {
                boundingBox.transform.Parent = parent.transform;
            }

            isFirstInParagraph = run.IsFirstInParagraph;

            Measurer = new FontMeasurerTrueType(measurementFont);
            fontStyles = measurementFont.Style;

            fontSizeInPixels = ((double)measurementFont.Size).PointToPixel(true);

            horizontalTextAlignment = run.Paragraph.HorizontalAlignment;

            if (run.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + run.Fill.Color.To6CharHexString();
            }

            Lines = GetLines(originalText);
        }

        internal List<string> GetLines(string text)
        {
            var inputWidth = WrapText ? MaxWidthPixels : double.NaN;
            Lines = TextWrapper.GetLines(text, Measurer, inputWidth);
            return Lines;
        }

        private int GetNumberOfLines()
        {
            return Lines.Count();
        }

        internal void InsertLineBreak(int insertPosition)
        {
            currentText = currentText.Insert(insertPosition, Environment.NewLine);
        }


        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = boundingBox.X;
            it = boundingBox.Y;

            ir = CalculateRightPositionInPixels();
            ib = CalculateBottomPositionInPixels();

            boundingBox.Right = ir;
            boundingBox.Bottom = ib;
        }

        internal double CalculateRightPositionInPixels()
        {
            var numLines = GetNumberOfLines();
            double retPos = boundingBox.X;

            double TextLengthInPixels = 0;

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
            var lineSize = ((double)measurementFont.Size).PointToPixel(true) + boundingBox.Y;
            var bottomPosition = lineSize * GetNumberOfLines();
            return bottomPosition;
        }

        internal double CalculateTextWidth(string targetString)
        {
            Measurer.MeasureWrappedTextCells = true;
            var width = Measurer.MeasureTextWidth(targetString);

            return width;
        }

        internal void CreateLines()
        {
            //string finalString = "";

            //Lines = SplitIntoLines(currentText);

            //foreach (var line in Lines)
            //{
            //    //    //Despite new textrun it could still be on the same line as previous textrun
            //    //    //Therefore only do line increase if we are first in paragraph or if we are not Lines[0].
            //    //    //This as line == Lines[0] && isFirstInParagraph == false means we are continuing on the same line as previous textRun
            //    //    //This is important if for example we have rich text where two letters on the same line has different colors.
            //    //    if (line != Lines[0] | isFirstInParagraph)
            //    //    {
            //    //        var yIncrease = /*isFirstInParagraph && useBaselineSpacing ? BaselineSpacing : LineSpacingPerNewLine;*/
            //    //        isFirstInParagraph = false;

            //    //    //    yIncrease = EPPlus.Fonts.OpenType.Utils.TextUtils.RoundToWhole(yIncrease);

            //    //    //    _yEndPos += yIncrease;
            //    //    //    if (Double.IsNaN(ClippingHeight) == false && _yEndPos >= ClippingHeight)
            //    //    //    {
            //    //    //        visibility = "display=\"none\"";
            //    //    //    }

            //    //    //    var yIncreaseString = yIncrease.ToString(CultureInfo.InvariantCulture);
            //    //    //    var xString = $"x =\"{(_xPosition).ToString(CultureInfo.InvariantCulture)}\" ";
            //    //    //    var dyString = $"dy =\"{yIncreaseString}px\" ";
            //    //    //    finalString += xString;
            //    //    //    finalString += dyString;
            //    //    //}

            //    //    //finalString += $"{visibility} " + $"{fontStyleAttributes} ";
            //    //    //if (measurementFont != null)
            //    //    //{
            //    //    //    finalString += $" font-family=\"{measurementFont.FontFamily},"
            //    //    //        + $"{measurementFont.FontFamily}_MSFontService,sans-serif\" "
            //    //    //        + $"font-size=\"{fontSizeInPixels.ToString(CultureInfo.InvariantCulture)}px\" ";
            //    //    //}
            //    //    //sb.Append(finalString);
            //    //    ////Get color etc.
            //    //    //base.Render(sb);
            //    //    //finalString = "";

            //    //    //finalString += ">";
            //    //    //finalString += line;
            //    //    //finalString += "</tspan>";
            //    //}
            //}
        }
    }
}
