using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems
{
    internal abstract class TextRunRenderItem : RenderItem
    {
        /// <summary>
        /// BoundingBox
        /// </summary>
        internal Rect BoundingBox;
        public override SvgItemType Type => SvgItemType.Text;

        internal readonly string originalText;
        private string currentText;

        MeasurementFont measurementFont;
        bool isFirstInParagraph;
        FontMeasurerTrueType textMeasurer;
        eTextAlignment horizontalTextAlignment;
        MeasurementFontStyles fontStyles;
        double fontSizeInPixels;

        List<string> Lines;

        internal TextRunRenderItem(ExcelParagraphTextRunBase run) 
        {
            originalText = run.Text;
            currentText = originalText;
            Lines = SplitIntoLines(originalText);

            BoundingBox = new Rect();

            isFirstInParagraph = run.IsFirstInParagraph;

            textMeasurer = new FontMeasurerTrueType(measurementFont);
            fontStyles = measurementFont.Style;

            fontSizeInPixels = ((double)measurementFont.Size).PointToPixel(true);

            horizontalTextAlignment = run.Paragraph.HorizontalAlignment;

            if (run.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + run.Fill.Color.To6CharHexString();
            }
        }

        internal List<string> SplitIntoLines(string text)
        {
            return text.Split(new string[] { "\r\n" }, StringSplitOptions.None).ToList();
        }
        private int GetNumberOfLines()
        {
            if (originalText != null)
            {
                if (currentText == null)
                {
                    currentText = originalText;
                }

                SplitIntoLines(currentText);
                return Lines.Count;
            }
            else
            {
                return 0;
            }
        }

        internal void InsertLineBreak(int insertPosition)
        {
            currentText = currentText.Insert(insertPosition, Environment.NewLine);
        }


        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = BoundingBox.X;
            it = BoundingBox.Y;

            //ir = CalculateRightPositionInPixels();
            //ib = CalculateBottomPositionInPixels();

            //BoundingBox.Left = il;
            //BoundingBox.Top = it;
            //BoundingBox.Right = ir;
            //BoundingBox.Bottom = ib;
        }

        internal double CalculateRightPositionInPixels()
        {
            var numLines = GetNumberOfLines();
            double retPos = BoundingBox.X;

            var TextLengthInPixels = 0;

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
            textMesurer.MeasureWrappedTextCells = true;
            var width = textMesurer.MeasureTextWidth(targetString);

            return width;
        }
    }
}
