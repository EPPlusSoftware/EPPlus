using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems
{
    internal class TextRunRenderItem : RenderItem
    {
        public override RenderItemType Type => RenderItemType.Text;

        internal readonly string _originalText;
        private string _currentText;

        MeasurementFont _measurementFont;
        bool _isFirstInParagraph;
        FontMeasurerTrueType _measurer;
        eTextAlignment _horizontalTextAlignment;
        MeasurementFontStyles _fontStyles;
        internal double FontSizeInPixels { get; private set; }

        List<string> Lines;

        double _yEndPos;
        double BaselineSpacing;

        bool _wrapText;
        double _maxWidthPixels = double.NaN;

        bool _isItalic = false;
        bool _isBold = false;
        eUnderLineType _underLineType;
        eStrikeType _strikeType;
        Color _underlineColor;

        internal double LineSpacingPerNewLine { get; set; }
        internal double BaseLineSpacing { get; set; }

        internal TextRunRenderItem(ExcelParagraphTextRunBase run, BoundingBox parent = null, string displayText = "")
        {
            _originalText = run.Text;
            _currentText = string.IsNullOrEmpty(displayText) ? _originalText : displayText;

            Lines = _currentText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();

            var measurer = run.Paragraph._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
            _measurer = (FontMeasurerTrueType)measurer;

            _measurementFont = run.GetMeasurementFont();
            _measurer.SetFont(_measurementFont);

            _isFirstInParagraph = run.IsFirstInParagraph;

            if (parent != null)
            {
                Bounds.Parent = parent;
            }

            _fontStyles = _measurementFont.Style;

            FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);

            _horizontalTextAlignment = run.Paragraph.HorizontalAlignment;

            if (run.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + run.Fill.Color.To6CharHexString();
            }

            _isItalic = run.FontItalic;
            _isBold = run.FontBold;
            _underLineType = run.FontUnderLine;
            _underlineColor = run.UnderLineColor;
            _strikeType = run.FontStrike;
        }

        internal List<string> GetLines(string text)
        {
            //var inputWidth = _wrapText ? _maxWidthPixels : double.NaN;
            //Lines = TextWrapper.GetLines(text, _measurer, inputWidth);
            return text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList(); ;
        }

        private int GetNumberOfLines()
        {
            return Lines.Count();
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.X;
            it = Bounds.Y;

            ir = CalculateRightPositionInPixels();
            ib = CalculateBottomPositionInPixels();

            Bounds.Right = ir;
            Bounds.Bottom = ib;
        }

        internal double CalculateRightPositionInPixels()
        {
            var numLines = GetNumberOfLines();
            double retPos = Bounds.X;

            double TextLengthInPixels = 0;

            if (numLines <= 0 || string.IsNullOrEmpty(_originalText))
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
            var lineSize = ((double)_measurementFont.Size).PointToPixel(true) + Bounds.Y;
            var bottomPosition = lineSize * GetNumberOfLines();
            return bottomPosition;
        }

        internal double CalculateTextWidth(string targetString)
        {
            _measurer.MeasureWrappedTextCells = true;
            var width = _measurer.MeasureTextWidth(targetString);

            return width;
        }

        public override void Render(StringBuilder sb)
        {
            throw new NotImplementedException();
        }

        //internal void CalculateTextWrapping(double maxWidth)
        //{
        //    List<string> NewContentLines = new List<string>();
        //    _measurer.SetFont(_measurementFont);
        //    var newLines = _measurer.MeasureAndWrapText(originalText, _measurementFont, maxWidth);
        //    Lines = newLines;
        //}

        //internal void CreateLines()
        //{
        //    //string finalString = "";

        //    //Lines = SplitIntoLines(currentText);

        //    //foreach (var line in Lines)
        //    //{
        //    //    //    //Despite new textrun it could still be on the same line as previous textrun
        //    //    //    //Therefore only do line increase if we are first in paragraph or if we are not Lines[0].
        //    //    //    //This as line == Lines[0] && isFirstInParagraph == false means we are continuing on the same line as previous textRun
        //    //    //    //This is important if for example we have rich text where two letters on the same line has different colors.
        //    //    //    if (line != Lines[0] | isFirstInParagraph)
        //    //    //    {
        //    //    //        var yIncrease = /*isFirstInParagraph && useBaselineSpacing ? BaselineSpacing : LineSpacingPerNewLine;*/
        //    //    //        isFirstInParagraph = false;

        //    //    //    //    yIncrease = EPPlus.Fonts.OpenType.Utils.TextUtils.RoundToWhole(yIncrease);

        //    //    //    //    _yEndPos += yIncrease;
        //    //    //    //    if (Double.IsNaN(ClippingHeight) == false && _yEndPos >= ClippingHeight)
        //    //    //    //    {
        //    //    //    //        visibility = "display=\"none\"";
        //    //    //    //    }

        //    //    //    //    var yIncreaseString = yIncrease.ToString(CultureInfo.InvariantCulture);
        //    //    //    //    var xString = $"x =\"{(_xPosition).ToString(CultureInfo.InvariantCulture)}\" ";
        //    //    //    //    var dyString = $"dy =\"{yIncreaseString}px\" ";
        //    //    //    //    finalString += xString;
        //    //    //    //    finalString += dyString;
        //    //    //    //}

        //    //    //    //finalString += $"{visibility} " + $"{fontStyleAttributes} ";
        //    //    //    //if (measurementFont != null)
        //    //    //    //{
        //    //    //    //    finalString += $" font-family=\"{measurementFont.FontFamily},"
        //    //    //    //        + $"{measurementFont.FontFamily}_MSFontService,sans-serif\" "
        //    //    //    //        + $"font-size=\"{fontSizeInPixels.ToString(CultureInfo.InvariantCulture)}px\" ";
        //    //    //    //}
        //    //    //    //sb.Append(finalString);
        //    //    //    ////Get color etc.
        //    //    //    //base.Render(sb);
        //    //    //    //finalString = "";

        //    //    //    //finalString += ">";
        //    //    //    //finalString += line;
        //    //    //    //finalString += "</tspan>";
        //    //    //}
        //    //}
        //}
    }
}
