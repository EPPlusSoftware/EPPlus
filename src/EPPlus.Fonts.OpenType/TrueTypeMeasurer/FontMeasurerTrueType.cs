/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using OfficeOpenXml.Interfaces.Fonts;

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

        OpenTypeFont LastFont = null;

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
            SetFont(mFont);
        }


        public void SetFont(double fontSize, string fontName, FontSubFamily subFamily = FontSubFamily.Regular)
        {
            CurrentFontName = fontName;
            var newFont = TextData.GetFontData(fontName, subFamily);
            CurrentFont = newFont;
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
        /// <summary>
        /// ASCENT is in shapes in Excel the distance between
        /// The top of the shape (or more likely for non-rect shapes the top of the inset rect textbox) 
        /// and the baseline of the given text.
        /// </summary>
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

        public List<TextLineSimple> MeasureAndWrapTextLines(string text, MeasurementFont font, double maxWidthPoints, double preExistingWidthPixels = 0)
        {
            SetFont(font.Size, font.FontFamily);
            var wrappedStrings = TextData.MeasureAndWrapTextLines(text, FontSize, CurrentFont, maxWidthPoints, preExistingWidthPixels.PixelToPoint());
            return wrappedStrings;
        }

        public List<TextLineSimple> MeasureAndWrapTextLines_New(string text, MeasurementFont font, double maxWidthPoints, double preExistingWidthPixels = 0)
        {
            List<TextLineSimple> lines = null;
            if (string.IsNullOrEmpty(text) == false)
            {
                var layout = TextData.GetTextLayoutEngine(font);
                var txtFragment = new Integration.TextFragment()
                {
                    Font = font,
                    Text = text
                };

                var list = new List<Integration.TextFragment>() { txtFragment };
                lines = layout.WrapRichTextLines(list, maxWidthPoints);
            }

            return lines;
            //SetFont(font.Size, font.FontFamily);
            //var wrappedStrings = TextData.MeasureAndWrapTextLines(text, FontSize, CurrentFont, maxWidthPoints, preExistingWidthPixels.PixelToPoint());
            //return wrappedStrings;
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

        public List<double> GetGlyphAdvanceWidths(string text, MeasurementFont font)
        {
            var fontdata = TextData.GetFontData(font.FontFamily, GetFontSubType(font.Style));
            var glyphdata = TextData.GetBoundsOfEachGlyph(text, fontdata);
            List<double> widths = new List<double>();
            foreach(var glyph in glyphdata)
            {
                widths.Add(glyph.advanceWidth * font.Size);
            }
            return widths;
        }

        public List<string> WrapMultipleTextFragments(List<string> textFragments, List<MeasurementFont> fonts, double maxWidthPoints)
        {
            TextFragmentCollection fragments = new TextFragmentCollection(textFragments);
            TextParagraph paragraph = new TextParagraph(fragments, fonts);

            //Wrap the fragments
            return TextData.WrapMultipleTextFragments(paragraph, maxWidthPoints);
        }

        public List<string> WrapMultipleTextFragments(TextFragmentCollection fragments, List<MeasurementFont> fonts, double maxWidthPoints)
        {
            TextParagraph paragraph = new TextParagraph(fragments, fonts);

            //Wrap the fragments
            return TextData.WrapMultipleTextFragments(paragraph, maxWidthPoints);
        }

        public List<TextLineSimple> WrapMultipleTextFragmentsToTextLines(TextFragmentCollection fragments, List<MeasurementFont> fonts, double maxWidthPoints)
        {
            TextParagraph paragraph = new TextParagraph(fragments, fonts);

            //Wrap the fragments
            return TextData.WrapMultipleTextFragmentsToTextLines(paragraph, maxWidthPoints);
        }

        public List<TextLineSimple> WrapMultipleTextFragmentsToTextLines_New(List<Integration.TextFragment> frags, double maxWidthPoints)
        {
            var layout = TextData.GetTextLayoutEngine(frags[0].Font);
            var lines = layout.WrapRichTextLines(frags, maxWidthPoints);
            return lines;
        }

        public List<TextLineSimple> WrapMultipleTextFragmentsToTextLines_New(List<string> rtTextFrags, List<MeasurementFont> fonts, double maxWidthPoints)
        {
            var layout = TextData.GetTextLayoutEngine(fonts[0]);

            var newFrags = new List<Integration.TextFragment>();

            for (int i = 0; i < rtTextFrags.Count(); i++)
            {
                var currentFrag = new Integration.TextFragment() { Text = rtTextFrags[i], Font = fonts[i] };
                newFrags.Add(currentFrag);
            }

            return layout.WrapRichTextLines(newFrags, maxWidthPoints);
            //TextParagraph paragraph = new TextParagraph(fragments, fonts);

            ////Wrap the fragments
            //return TextData.WrapMultipleTextFragmentsToTextLines(paragraph, maxWidthPoints);
        }
    }
}
