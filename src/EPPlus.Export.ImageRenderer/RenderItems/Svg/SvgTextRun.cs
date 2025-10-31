using EPPlusImageRenderer.FontData;
using EPPlusImageRenderer.Svg;
using EPPlusImageRenderer.Text;
using EPPlusImageRenderer.Utils;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer.RenderItems.Svg
{
    internal class SvgTextRun : SvgRenderItem
    {
        double LineSpacingPerNewLine;
        double yPosition;
        double yClipHeight;
        double fontSizeInPixels;
        eTextAlignment horizontalTextAlignment;
        string textContent;

        internal RectBase BoundingBox = new();
        internal Coordinate origin = new Coordinate(0, 0);
        internal ExcelParagraphTextRunBase TextRun;
        List<string> Lines;

        MeasurementFont measurementFont;

        FontMeasurerExact fmExact;

        //Unnecesary??
        double TextLengthInPixels;

        string horizontalAttribute;

        /// <summary>
        /// constructor. enter linespacing in pixels
        /// </summary>
        /// <param name="textRun"></param>
        /// <param name="lineSpacing"></param>
        internal SvgTextRun(ExcelParagraphTextRunBase textRun, double lineSpacing, double textMaxWidth) : base()
        {
            TextRun = textRun;
            Lines = SplitIntoLines();

            measurementFont = TextRun.GetMeasureFont();

            fmExact = new FontMeasurerExact(measurementFont);

            horizontalTextAlignment = TextRun.Paragraph.HorizontalAlignment;

            LineSpacingPerNewLine = lineSpacing;

            //No idea how to get this property now since we have no access to textbody
            bool wrapText = true;
            if (wrapText)
            {
                CalculateTextWrapping(textMaxWidth);
            }
        }
        public RectBase textArea;

        public override SvgItemType Type => SvgItemType.TSpan;

        public override void Render(StringBuilder sb)
        {
            string finalString = "";

            foreach (var line in Lines)
            {
                yPosition += TextUtils.RoundToWhole(LineSpacingPerNewLine);
                string visibility = "";
                if (yPosition >= yClipHeight)
                {
                    visibility = "display=\"none\"";
                }

                finalString += $"<tspan {visibility} font-size=\"{(fontSizeInPixels / 16).ToString(CultureInfo.InvariantCulture)}em\" dy=\"{LineSpacingPerNewLine.ToString(CultureInfo.InvariantCulture)}px\">";
                finalString += line;
                finalString += "</tspan>";
            }

            sb.Append(finalString);
            //throw new NotImplementedException();
        }

        internal override RenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }

        internal double GetNumberOfLines()
        {
            if (TextRun != null)
            {
                SplitIntoLines();
                return Lines.Count;
            }
            else
            {
                return 0;
            }
        }

        internal List<string> SplitIntoLines()
        {
            return TextRun.Text.Split("\r\n").ToList();
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

            if (numLines <= 0 || string.IsNullOrEmpty(TextRun.Text))
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
            var bottomPosition = (((double)TextRun.FontSize).PointToPixel() + origin.Y) * GetNumberOfLines();
            return bottomPosition;
        }

        internal double CalculateTextWidth(string targetString)
        {
            var textMesurer = new FontMeasurerExact(measurementFont);
            textMesurer.MeasureWrappedTextCells = true;
            var width = textMesurer.MeasureTextWidth(targetString);

            return width;
        }

        internal void CalculateTextWrapping(double maxWidth)
        {
            List<string> NewContentLines = new List<string>();
            var textMesurer = new FontMeasurerExact(measurementFont);

            var newLines = textMesurer.MeasureAndWrapText(TextRun.Text, measurementFont, maxWidth);
            Lines = newLines;
        }
    }
}
