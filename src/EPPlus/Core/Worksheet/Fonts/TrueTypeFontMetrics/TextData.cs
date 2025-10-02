using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Cmap;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Kern;
using System.Collections.Generic;
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader;
using System;
using System.Linq;

namespace EPPlusImageRenderer.TrueTypeFontMeasurer
{
    internal static class TextData
    {
        /// <summary>
        /// Add additional folders to search for fonts.
        /// </summary>
        public static List<string> FontDirectories = new List<string>();

        /// <summary>
        /// If true, epplus will look for fonts in the system directories for installed fonts.  c:\windows\fonts
        /// </summary>
        public static bool SearchSystemDirectories = true;

        internal static TtfFont GetFontData(string fontName, string subFamily)
        {
            return GenericFonts.GetFontData(FontDirectories, SearchSystemDirectories, fontName, subFamily);
        }

        /// <summary>
        /// Calculates the height above the baseline minus the height below the baseline
        /// Returning the resulting middle line or x-height of the font
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double MeasureFontHeight(TtfFont font, double fontSize)
        {
            var asc = font.Os2Table.usWinAscent;
            var desc = font.Os2Table.usWinDescent;
            var size = fontSize;
            var em = font.HeadTable.UnitsPerEm;
            var lineHeight = asc + desc;
            var lineHeightPt = lineHeight * (size / em);
            return lineHeightPt;
        }

        /// <summary>
        /// ASCENT is in shapes in Excel the distance between
        /// The top of the shape (or more likely for non-rect shapes the top of the inset rect textbox) 
        /// and the baseline of the given text.
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double GetBaseLine(TtfFont font, double fontSize)
        {
            if(font.Os2Table.UseTypoMetrics)
            {
                var typoAscent = font.Os2Table.sTypoAscender;
                var em = font.HeadTable.UnitsPerEm;

                //var pixelSize = (typoAscent * fontSize * 96d) / (72d * em);

                return typoAscent * (fontSize / em);
            }
            else
            {
                var asc = font.Os2Table.usWinAscent;
                var em = font.HeadTable.UnitsPerEm;
                return asc * (fontSize / em);
            }
        }

        /// <summary>
        /// TODO: Support HHead for MAC line spacing
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double GetSingleLineSpacing(TtfFont font, double fontSize)
        {
            var singleLineSpacing = font.Os2Table.UseTypoMetrics ? MeasureSingleLineSpacing_sTypo(font, fontSize) : MeasureFontHeight(font, fontSize);

            return singleLineSpacing;
        }

        /// <summary>
        /// In modern fonts 
        /// USE_TYPO_METRICS is always checked.
        /// If it is this method is used to determine line-spacing.
        /// If not usWinAscent and usWinDescent is used for windows
        /// HHead ascent and descent for mac.
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double MeasureSingleLineSpacing_sTypo(TtfFont font, double fontSize)
        {
            var typoAscent = font.Os2Table.sTypoAscender;
            var typoDescent = font.Os2Table.sTypoDescender;
            var typoLineGap = font.Os2Table.sTypoLineGap;

            double em = font.HeadTable.UnitsPerEm;
            //In "Typo" attributes descent is always negative if lower than the baseline (which it is for a vast majority of fonts
            //Calculate baseline to baseline line spacing. See: https://learn.microsoft.com/en-us/typography/opentype/spec/recom#tad
            double lineHeight = typoAscent - typoDescent + typoLineGap;
            // var lineHeightPt = lineHeight * (fontSize / em);

            //Translate from font design units to points
            var lineHeightPt = (lineHeight / em) * fontSize;
            return lineHeightPt;
        }

        /// <summary>
        /// Measures distance between Top and Bottom typography lines beyond ascent and descent
        /// (Includes subscript and superscript heights?)
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double MeasureBoundingBoxHeight(TtfFont font, double fontSize)
        {
            var max = font.HeadTable.Ymax;
            var min = font.HeadTable.Ymin;

            var em = font.HeadTable.UnitsPerEm;

            var height = max - min;

            var heightInPoints = ((double)height / (double)em) * fontSize;
            return heightInPoints;
        }

        /// <summary>
        /// Measures largest possible glyph width
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double MeasureBoundingBoxWidth(TtfFont font, double fontSize)
        {
            var max = font.HeadTable.Xmax;
            var min = font.HeadTable.Xmin;

            var em = font.HeadTable.UnitsPerEm;

            var width = max - min;

            var widthPt = width * (fontSize / em);
            return widthPt;
        }

        internal static double MeasureAscent(TtfFont font, double fontSize)
        {
            return font.Os2Table.usWinAscent * (fontSize / font.HeadTable.UnitsPerEm);
        }

        internal static double MeasureDescent(TtfFont font, double fontSize)
        {
            return font.Os2Table.usWinDescent * (fontSize / font.HeadTable.UnitsPerEm);
        }

        /// <summary>
        /// Measures the text and breaks it into smaller strings so that none exceed the MaxWidth
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontData"></param>
        /// <param name="maxWidth"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        internal static List<string> MeasureAndWrapText(string text, double fontSize, TtfFont fontData, double maxWidth)
        {
            double totalAdvanceWidth = 0;
            ushort lastGlyphIndex = 0;
            bool firstChar = true;

            List<string> wrappedStrings = new List<string>();
            int previousLineIndex = 0;

            //Convert maxWidth from points to font design units
            maxWidth = (maxWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                var encodingRecord = fontData.CmapTable.EncodingRecords.FirstOrDefault(er => er.PlatformId == Platforms.Windows && er.EncodingId == 1);
                if (encodingRecord == null) throw new Exception("Could not find Microsoft Unicode cmap (PlatformID 3, EncodingID 1).");
                GlyphMapping[] mappings = encodingRecord.Mappings;
                encodingRecord.CharMappingsToGlyphIndex.TryGetValue(c, out ushort gi);
                int advanceWidth;
                if (gi == 0 && c != 0)
                {
                    advanceWidth = fontData.Os2Table.xAvgCharWidth;
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi];
                    advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                }

                var newWidth = totalAdvanceWidth + advanceWidth;

                // Kerning adjustment
                if (!firstChar)
                {
                    int kerning = GetKerningAdjustment(lastGlyphIndex, gi, fontData);
                    newWidth += kerning;
                }

                if (newWidth > maxWidth)
                {
                    var txt = text.Substring(previousLineIndex, i - previousLineIndex);

                    //Ensure whole words get moved down if part of its letters are overflowing
                    var splitLines = txt.Split(' ');

                    if (splitLines.Length > 1 && c != ' ')
                    {
                        var stringOverMax = splitLines.Last();
                        var startIndex = txt.Length - stringOverMax.Length;
                        //Remove the overflowing characters
                        var spacedString = txt.Remove(startIndex, stringOverMax.Length).TrimEnd(' ');
                        //Add only part of the text before the overflowing word
                        wrappedStrings.Add(spacedString);

                        //Calculate the new line index
                        previousLineIndex = i - stringOverMax.Length;

                        totalAdvanceWidth = advanceWidth;
                    }
                    else
                    {
                        wrappedStrings.Add(txt.Remove(i));
                        previousLineIndex = c == ' ' ? i + 1 : i;
                        totalAdvanceWidth = 0;
                    }
                }
                else
                {
                    totalAdvanceWidth = newWidth;
                }

                lastGlyphIndex = gi;
                firstChar = false;
            }

            wrappedStrings.Add(text.Substring(previousLineIndex));

            return wrappedStrings;
        }

        /// <summary>
        /// If wrap text is true returns the largest width before newLine characters
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontData"></param>
        /// <param name="wrapText"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        internal static double MeasureText(string text, double fontSize, TtfFont fontData, bool wrapText = false)
        {
            double totalAdvanceWidth = 0;
            ushort lastGlyphIndex = 0;
            bool firstChar = true;

            ////For if we want to calculate the total glyph height within a specific string
            //short GreatestYMax = short.MinValue;
            //short LowestYMin = short.MaxValue;

            double largestWidth = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                var encodingRecord = fontData.CmapTable.EncodingRecords.FirstOrDefault(er => er.PlatformId == Platforms.Windows && er.EncodingId == 1);
                if (encodingRecord == null) throw new Exception("Could not find Microsoft Unicode cmap (PlatformID 3, EncodingID 1).");
                GlyphMapping[] mappings = encodingRecord.Mappings;
                encodingRecord.CharMappingsToGlyphIndex.TryGetValue(c, out ushort gi);
                int advanceWidth;
                if (gi == 0 && c != 0)
                {
                    advanceWidth = fontData.Os2Table.xAvgCharWidth;
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi];
                    advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                }

                if (wrapText && (c == '\n' || c == '\r'))
                {
                    if (i > 0 && c == '\r' && text[i - 1] == '\n')
                    {
                        continue; //CRLF should be handle
                                  //d as one new line.
                    }
                    if (totalAdvanceWidth > largestWidth)
                    {
                        largestWidth = totalAdvanceWidth;
                        totalAdvanceWidth = 0;
                    }
                }

                totalAdvanceWidth += advanceWidth;
                // Kerning adjustment
                if (!firstChar)
                {
                    int kerning = GetKerningAdjustment(lastGlyphIndex, gi, fontData);
                    totalAdvanceWidth += kerning;
                }

                ////For if we want to calculate the total glyph height within a specific string
                //var yMax = fontData.GlyphTable.Glyphs[gi].yMax;
                //var yMin = fontData.GlyphTable.Glyphs[gi].yMin;

                //if(yMax > GreatestYMax)
                //{
                //    GreatestYMax = yMax;
                //}

                //if(yMin < LowestYMin)
                //{
                //    LowestYMin = yMin;
                //}

                lastGlyphIndex = gi;
                firstChar = false;
            }

            largestWidth =  largestWidth < totalAdvanceWidth ? totalAdvanceWidth : largestWidth;

            ////For if we want to calculate the total glyph height within a specific string
            //var height = GreatestYMax - LowestYMin;
            //var em = fontData.HeadTable.UnitsPerEm;
            //var heightPt = height * (fontSize / em);

            // Convert to points
            return (largestWidth / (double)fontData.HeadTable.UnitsPerEm) * fontSize;
        }

        private static int GetKerningAdjustment(ushort left, ushort right, TtfFont fontData)
        {
            foreach (var subtable in fontData.KernTable.SubTables)
            {
                if (subtable.Format0Subtable == null) continue;
                // Format 0 only
                int format = subtable.coverage._coverage >> 8;
                bool isHorizontal = (subtable.coverage._coverage & 0x1) == 1;
                if (format != 0 || !isHorizontal) continue;
                KerningPair[] pairs = subtable.Format0Subtable.Pairs;
                if (pairs == null) continue;
                for (int i = 0; i < pairs.Length; i++)
                {
                    if (pairs[i].left == left && pairs[i].right == right)
                        return pairs[i].value;
                }
            }
            return 0;
        }
    }
}
