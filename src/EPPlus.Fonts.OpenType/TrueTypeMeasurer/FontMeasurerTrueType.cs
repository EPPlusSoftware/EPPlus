using System.Collections.Generic;
using OfficeOpenXml.Interfaces.Drawing.Text;
using EPPlus.Fonts.OpenType.Utils;
using System;

namespace EPPlus.Fonts.OpenType
{
    public enum TextUnit
    {
        Points,
        Pixels
    }

    public class FontMeasurerTrueType : ITextMeasurerWrap
    {
        private OpenTypeFont _defaultFont = TextData.GetFontData("Calibri", FontSubFamily.Regular);
        private double _defaultSize = 11d;

        OpenTypeFont CurrentFont;
        private double _widthInPoints;
        private double _heightInPoints;

        /// <summary>
        /// FontSize can be seen but must be set via SetFont method
        /// </summary>
        internal double FontSize { get; private set; }
        /// <summary>
        /// FontName can be seen but must be set via SetFont method
        /// </summary>
        internal string CurrentFontName { get; private set; }

        public bool MeasureWrappedTextCells { get; set; } = false;

        /// <summary>
        /// Creates a font measurer with the default font
        /// </summary>
        public FontMeasurerTrueType()
        {
            //Set only to avoid warning. Set font already sets it.
            CurrentFont = _defaultFont;

            CurrentFontName = _defaultFont.GetEnglishFontFamilyName();

            SetFont(_defaultSize, CurrentFontName);
        }

        /// <summary>
        /// FontSize is assumed to be in Points
        /// </summary>
        /// <param name="fontSize"></param>
        /// <param name="fontName"></param>
        /// <param name="fontSubFamily"></param>
        public FontMeasurerTrueType(double fontSize, string fontName, FontSubFamily fontSubFamily = FontSubFamily.Regular)
        {
            CurrentFontName = string.IsNullOrEmpty(fontName) ? "" : fontName;

            SetFont(fontSize, fontName, fontSubFamily);

            if (CurrentFont == null)
            {
                CurrentFont = _defaultFont;
            }
        }

        public FontMeasurerTrueType(MeasurementFont mFont)
        {
            CurrentFontName = string.IsNullOrEmpty(mFont.FontFamily) ? "" : mFont.FontFamily;

            SetFont(mFont.Size, mFont.FontFamily);

            if (CurrentFont == null)
            {
                CurrentFont = _defaultFont;
            }
        }


        public void SetFont(double fontSize, string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            CurrentFontName = fontName;
            CurrentFont = TextData.GetFontData(fontName, subFamily);
            FontSize = fontSize;
        }

        public void SetFont(MeasurementFont mFont)
        {
            CurrentFontName = string.IsNullOrEmpty(mFont.FontFamily) ? "" : mFont.FontFamily;

            var styleName = GetFontSubType(mFont.Style);
            SetFont(mFont.Size, mFont.FontFamily, styleName);

            if (CurrentFont == null)
            {
                CurrentFont = _defaultFont;
            }
        }

        /// <summary>
        /// Returns exact textWidth in points
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public double MeasureTextWidth(string text)
        {
            _widthInPoints = TextData.MeasureText(text, FontSize, CurrentFont, MeasureWrappedTextCells);

            return _widthInPoints;
        }
        /// <summary>
        /// Returns exact textWidth in pixels
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public double MeasureTextWidthInPixels(string text)
        {
            MeasureTextWidth(text);
            return _widthInPoints.PointToPixel();
        }

        public double GetMinXInPixels()
        {
            return ((Math.Abs((double)167)/2048d) * FontSize).PointToPixel();
        }

        /// <summary>
        /// Gets the height of a bounding box that can contain every glyph in the font.
        /// Nothing should be able to go beyond this height unless we've missed some edge-case.
        /// </summary>
        /// <returns></returns>
        public double GetLargestPossibleHeight()
        {
            _heightInPoints = TextData.MeasureBoundingBoxHeight(CurrentFont, FontSize);
            return _heightInPoints;
        }

        /// <summary>
        /// Gets and converts the height to pixels
        /// </summary>
        /// <returns></returns>
        public double GetLargestPossibleHeightInPixels()
        {
            GetLargestPossibleHeight();
            return _heightInPoints.PointToPixel();
        }

        public double InternalFontHeight()
        {
            return TextData.MeasureFontHeight(CurrentFont, FontSize);
        }


        public double InternalFontHeightPixel()
        {
            return InternalFontHeight().PointToPixel();
        }

        public double GetSingleLineSpacing()
        {
            return TextData.GetSingleLineSpacing(CurrentFont, FontSize);
        }

        public double GetBaseLine()
        {
            return TextData.GetBaseLine(CurrentFont, FontSize);
        }

        //bool? _validForEnvironment = null;
        public bool ValidForEnvironment()
        {
            return true;
        }

        /// <summary>
        /// Measures text width and height in points
        /// </summary>
        /// <param name="text"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public TextMeasurement MeasureText(string text, MeasurementFont font)
        {
            var TextMeasurement = new TextMeasurement();
            SetFont(font.Size, font.FontFamily, GetFontSubType(font.Style));
            TextMeasurement.Width = (float)MeasureTextWidth(text);
            TextMeasurement.Height = (float)GetSingleLineSpacing();
            TextMeasurement.FontHeight = (float)InternalFontHeight();
            return TextMeasurement;
        }

        private FontSubFamily GetFontSubType(MeasurementFontStyles Style)
        {
            if ((Style & (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic)) == (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic))
            {
                return FontSubFamily.BoldItalic;
            }
            else if ((Style & MeasurementFontStyles.Bold) == MeasurementFontStyles.Bold)
            {
                return FontSubFamily.Bold;
            }
            else if ((Style & MeasurementFontStyles.Italic) == MeasurementFontStyles.Italic)
            {
                return FontSubFamily.Italic;
            }

            return FontSubFamily.Regular;
        }
        
        /// <summary>
        /// Measures text width and height in points
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public TextMeasurement MeasureText(string text)
        {
            var TextMeasurement = new TextMeasurement();
            TextMeasurement.Width = (float)MeasureTextWidth(text);
            TextMeasurement.Height = (float)GetSingleLineSpacing();
            return TextMeasurement;
        }

        public List<string> MeasureAndWrapText(string text, MeasurementFont font, double MaxWidthInPixels, double preExistingWidthPixels = 0)
        {
            SetFont(font.Size, font.FontFamily);
            var wrappedStrings = TextData.MeasureAndWrapText(text, FontSize, CurrentFont, MaxWidthInPixels.PixelToPoint(), preExistingWidthPixels.PixelToPoint());
            return wrappedStrings;
        }

        public List<string> MeasureAndWrapTextPoints(string text, MeasurementFont font, double MaxWidthInPoints)
        {
            SetFont(font.Size, font.FontFamily);
            var wrappedStrings = TextData.MeasureAndWrapText(text, FontSize, CurrentFont, MaxWidthInPoints);
            return wrappedStrings;
        }

        public List<string> MeasureAndWrapText(string text, double MaxWidthInPixels)
        {
            var wrappedStrings = TextData.MeasureAndWrapText(text, FontSize, CurrentFont, MaxWidthInPixels.PixelToPoint());
            return wrappedStrings;
        }
        /// <summary>
        /// This assumes the first line is "Ascent only" as is usually the case in Paragraphs
        /// And that the last line has a full line of linespacing under it.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="firstLine"></param>
        /// <returns></returns>
        public double GetHeightOfText(List<string> text)
        {
            var distToBaselineFromTopOfContainer = GetBaseLine();
            var distBaselineToBaseline = GetSingleLineSpacing();

            var finalHeight = (distBaselineToBaseline * (text.Count));
            return finalHeight;
        }
        public double GetHeightOfTextInPixels(List<string> text)
        {
            return GetHeightOfText(text).PointToPixel();
        }

        public double GetDeltaAscent(TextUnit unit = TextUnit.Points)
        {
            var dAscent = TextData.GetDeltaAscent(CurrentFont, FontSize);
            if (unit == TextUnit.Pixels)
            {
                return dAscent.PointToPixel();
            }
            return dAscent;
        }

        public List<string> WrapMultipleTextFragments(List<string> textFragment, List<MeasurementFont> fonts, double maxWidthPoints)
        {
            List<OpenTypeFont> openTypeFonts = new List<OpenTypeFont>();
            List<double> fontSizes = new List<double>();
            foreach(var font in fonts)
            {
                SetFont(font);
                openTypeFonts.Add(CurrentFont);
                fontSizes.Add(font.Size);
            }

            return TextData.WrapMultipleTextFragments(textFragment, fontSizes, openTypeFonts, maxWidthPoints);
        }
    }
}
