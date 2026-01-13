using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class TextRunItem : SvgRenderItem
    {
        public override RenderItemType Type => RenderItemType.Text;

        internal readonly string _originalText;
        protected string _currentText;

        internal protected MeasurementFont _measurementFont;
        internal protected bool _isFirstInParagraph;
        FontMeasurerTrueType _measurer;
        eTextAlignment _horizontalTextAlignment;
        MeasurementFontStyles _fontStyles;
        internal double FontSizeInPixels { get; private set; }

        public List<string> Lines { get; private set; }

        double _yEndPos;
        double BaselineSpacing;

        bool _wrapText;
        double _maxWidthPixels = double.NaN;

        protected internal bool _isItalic = false;
        protected internal bool _isBold = false;
        protected internal eUnderLineType _underLineType;
        protected internal eStrikeType _strikeType;
        protected internal Color _underlineColor;

        internal double LineSpacingPerNewLine { get; set; }
        internal double BaseLineSpacing { get; set; }

        internal TextRunItem(ExcelParagraphTextRunBase run, BoundingBox parent = null, string displayText = "")
        {
            Bounds.transform.Name = "TextRun";

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
            return text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();
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
    }
}
