using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class TextRunItem : RenderItem
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

        /// <summary>
        /// Aka total Delta Y
        /// </summary>
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

        internal List<double> YIncreasePerLine { get; private set; } = new List<double>();
        internal List<double> PerLineWidth { get; private set; } = new List<double>();
        internal double ClippingHeight = double.NaN;

        internal TextRunItem(DrawingBase renderer, BoundingBox parent, string text, ExcelTextFont font, string displayText) : base(renderer, parent)
        {
            _originalText = text;

            Bounds.Name = "TextRun";
            _currentText = string.IsNullOrEmpty(displayText) ? _originalText : displayText;

            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            _measurer = new FontMeasurerTrueType();            

            _measurementFont = font.GetMeasureFont();
            _measurer.SetFont(_measurementFont);

            _isFirstInParagraph = true;

            _fontStyles = _measurementFont.Style;

            FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);
            Bounds.Height = FontSizeInPixels;
            if (parent.Height < FontSizeInPixels)
            {
                parent.Height = FontSizeInPixels;
            }
            _horizontalTextAlignment = eTextAlignment.Center;

            if (font.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + font.Fill.Color.To6CharHexString();
            }

            //To get clipping height we need to get the textbody bounds
            if (parent != null && parent.Parent != null && parent.Parent.Parent != null)
            {
                ClippingHeight = parent.Parent.Parent.LocalPosition.Y + parent.Parent.Parent.Size.Y;
            }
            if(Lines.Count==1)
            {
                Bounds.Width = parent.Width;
            }
            else
            {
                //Measure text.
            }
            _isItalic = font.Italic;
            _isBold = font.Bold;
            _underLineType = font.UnderLine;
            _underlineColor = font.UnderLineColor;
            _strikeType = font.Strike;
        }

        /// <summary>
        /// If the run has been wrapped more line-breaks may have been added in displayText
        /// </summary>
        /// <param name="run"></param>
        /// <param name="parent"></param>
        /// <param name="displayText"></param>
        internal TextRunItem(DrawingBase renderer, BoundingBox parent, ExcelParagraphTextRunBase run, string displayText = "") : base(renderer, parent)
        {
            _originalText = run.Text;

            Bounds.Name = "TextRun";
            _currentText = string.IsNullOrEmpty(displayText) ? _originalText : displayText;

            Lines = Regex.Split(_currentText, "\r\n|\r|\n").ToList();

            var measurer = run.Paragraph._prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
            _measurer = (FontMeasurerTrueType)measurer;

            _measurementFont = run.GetMeasurementFont();
            _measurer.SetFont(_measurementFont);

            _isFirstInParagraph = run.IsFirstInParagraph;

            _fontStyles = _measurementFont.Style;

            FontSizeInPixels = ((double)_measurementFont.Size).PointToPixel(true);

            _horizontalTextAlignment = run.Paragraph.HorizontalAlignment;

            if (run.Fill.IsEmpty == false && run.Fill.Style == eFillStyle.SolidFill)
            {
                FillColor = "#" + run.Fill.Color.To6CharHexString();
            }

            //To get clipping height we need to get the textbody bounds
            if( parent!= null && parent.Parent != null && parent.Parent.Parent != null)
            {
               ClippingHeight = ((BoundingBox)parent.Parent.Parent).Bottom;
            }

            _isItalic = run.FontItalic;
            _isBold = run.FontBold;
            _underLineType = run.FontUnderLine;
            _underlineColor = run.UnderLineColor;
            _strikeType = run.FontStrike;
        }

        internal double CalculateLineSpacing()
        {
            //---Calculation of total y-height---

            //If already did this
            if(YIncreasePerLine.Count > 0)
            {
                return Bounds.Height;
            }

            bool lineIsFirstInParagraph = _isFirstInParagraph;

            foreach (var line in Lines)
            {
                //Despite new textrun it could still be on the same line as previous textrun
                //Therefore only do line increase if we are first in paragraph or if we are not Lines[0].
                //This as line == Lines[0] && isFirstInParagraph == false means we are continuing on the same line as previous textRun
                //This is important if for example we have rich text where two letters on the same line has different colors.
                if (line != Lines[0] || lineIsFirstInParagraph)
                {
                    var yIncrease = lineIsFirstInParagraph ? BaseLineSpacing : LineSpacingPerNewLine;
                    lineIsFirstInParagraph = false;

                    //yIncrease = Fonts.OpenType.Utils.TextUtils.RoundToWhole(yIncrease);

                    YIncreasePerLine.Add(yIncrease);

                    _yEndPos += yIncrease;
                    //if (Double.IsNaN(ClippingHeight) == false && _yEndPos >= ClippingHeight)
                    //{
                    //    bool displayLine = false
                    //}
                }
                else
                {
                    YIncreasePerLine.Add(0);
                }
            }

            if(_yEndPos == 0)
            {
                Bounds.Height = FontSizeInPixels;
            }
            else
            {
                Bounds.Height = _yEndPos;
            }

            return _yEndPos;
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

        /// <summary>
        /// Calculates right/bottom
        /// </summary>
        /// <param name="il"></param>
        /// <param name="it"></param>
        /// <param name="ir"></param>
        /// <param name="ib"></param>
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.Left;
            it = Bounds.Top;

            //Sets bounds bottom correctly
            CalculateLineSpacing();
            ib = Bounds.Bottom;
            
            //Caculate bounds right
            ir = CalculateRightPositionInPixels();
            Bounds.Width = ir-Bounds.Left;
        }

        internal double CalculateRightPositionInPixels()
        {
            var numLines = GetNumberOfLines();
            double retPos = Bounds.Left;

            double textLengthPixels = 0;

            if (numLines <= 0 || string.IsNullOrEmpty(_originalText))
            {
                textLengthPixels = 0;
                return retPos;
            }

            double longestWidth = -1;

            for (int i = 0; i < numLines; i++)
            {
                var text = Lines[i];
                var width = CalculateTextWidth(Lines[i]).PointToPixel();
                PerLineWidth.Add(width);

                if (width > longestWidth)
                {
                    longestWidth = width;
                }
            }

            textLengthPixels = longestWidth;

            retPos += longestWidth;

            //Calculates right
            Bounds.Width = textLengthPixels;

            return Bounds.Right;
        }

        internal double CalculateBottomPositionInPixels()
        {
            return Bounds.Height;
        }

        internal double CalculateTextWidth(string targetString)
        {
            _measurer.SetFont(_measurementFont);
            _measurer.MeasureWrappedTextCells = true;
            var width = _measurer.MeasureTextWidth(targetString);

            return width;
        }
        //        origin.Left = x; origin.Top = y;
        //    origin.Left = x;
        //    origin.Top = y;
    }
}
