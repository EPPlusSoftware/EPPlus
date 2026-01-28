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
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace EPPlus.Fonts.OpenType
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

        internal static OpenTypeFont GetFontData(string fontName, FontSubFamily subFamily)
        {
            return OpenTypeFonts.GetFontData(FontDirectories, fontName, subFamily, SearchSystemDirectories);
        }

        /// <summary>
        /// Get difference between winAscent and typoAscent in points
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double GetDeltaAscent(OpenTypeFont font, double fontSize)
        {
            var winAscent = GetWinAscent(font, fontSize);
            var typoAscent = GetTypoAscent(font, fontSize);
            return winAscent - typoAscent;
        }

        /// <summary>
        /// Calculates the height above the baseline minus the height below the baseline
        /// Returning the resulting middle line or x-height of the font
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double MeasureFontHeight(OpenTypeFont font, double fontSize)
        {
            var asc = font.Os2Table.usWinAscent;
            var desc = font.Os2Table.usWinDescent;
            var size = fontSize;
            var em = font.HeadTable.UnitsPerEm;
            var lineHeight = asc + desc;
            var lineHeightPt = lineHeight * (size / em);
            return lineHeightPt;
        }

        internal static double GetTypoAscent(OpenTypeFont font, double fontSize)
        {
            var typoAscent = font.Os2Table.sTypoAscender;
            var em = font.HeadTable.UnitsPerEm;

            return typoAscent * (fontSize / em);
        }

        internal static double GetWinAscent(OpenTypeFont font, double fontSize)
        {
            var asc = font.Os2Table.usWinAscent;
            var em = font.HeadTable.UnitsPerEm;
            return asc * (fontSize / em);
        }

        /// <summary>
        /// ASCENT is in shapes in Excel the distance between
        /// The top of the shape (or more likely for non-rect shapes the top of the inset rect textbox) 
        /// and the baseline of the given text.
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double GetBaseLine(OpenTypeFont font, double fontSize)
        {
            if (font.Os2Table.UseTypoMetrics)
            {
                return GetTypoAscent(font, fontSize);
            }
            else
            {
                return GetWinAscent(font, fontSize);
            }
        }

        /// <summary>
        /// TODO: Support HHead for MAC line spacing
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double GetSingleLineSpacing(OpenTypeFont font, double fontSize)
        {
            var singleLineSpacing = font.Os2Table.UseTypoMetrics ? MeasureSingleLineSpacing_sTypo(font, fontSize) : MeasureFontHeight(font, fontSize);

            return singleLineSpacing;
        }
        internal static double GetSingleLineSpacing_new(OpenTypeFont font, double fontSize)
        {
            var singleLineSpacing = font.Os2Table.UseTypoMetrics ? MeasureSingleLineSpacing_sTypo(font, fontSize) : MeasureFontHeight_New(font, fontSize);

            return singleLineSpacing;
        }
        internal static double MeasureFontHeight_New(OpenTypeFont font, double fontSize)
        {
            var asc = font.HheaTable.ascender;
            var desc = font.HheaTable.descender;
            var lineGap = font.HheaTable.lineGap;
            var size = fontSize;
            var em = font.HeadTable.UnitsPerEm;
            var lineHeight = asc + desc + lineGap;
            var lineHeightPt = lineHeight * (size / em);
            return lineHeightPt;
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
        internal static double MeasureSingleLineSpacing_sTypo(OpenTypeFont font, double fontSize)
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
        internal static double MeasureBoundingBoxHeight(OpenTypeFont font, double fontSize)
        {
            var max = font.HeadTable.Ymax;
            var min = font.HeadTable.Ymin;

            var em = font.HeadTable.UnitsPerEm;

            var height = max - min;

            var heightInPoints = ((double)height / (double)em) * fontSize;
            return heightInPoints;
        }

        private static double MeasureBoundingBoxWidth(OpenTypeFont font)
        {
            var max = font.HeadTable.Xmax;
            var min = font.HeadTable.Xmin;

            var width = max - min;
            return width;
        }

        /// <summary>
        /// Measures largest possible glyph width
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <returns></returns>
        internal static double MeasureBoundingBoxWidth(OpenTypeFont font, double fontSize)
        {
            var width = MeasureBoundingBoxWidth(font);

            var em = font.HeadTable.UnitsPerEm;

            var widthPt = width * (fontSize / em);
            return widthPt;
        }


        /// <summary>
        /// Assumes string is pure string with no linebreaks or need for wrapping
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        internal static BoundingRectangle GetStringMaximumBoundingRectangle(string text, double fontSize, OpenTypeFont font)
        {
            var width = MeasureBoundingBoxWidth(font, fontSize);
            var height = MeasureBoundingBoxHeight(font, fontSize);

            var numChars = text.Count();
            var maxLineWidth = width * numChars;

            BoundingRectangle boundingBox = new BoundingRectangle() { Xmin = 0, Ymin = 0, Xmax = (short)maxLineWidth, Ymax = (short)height };
            return boundingBox;
        }

        /// <summary>
        /// Get bounds/widths for each glyph in font desugn units
        /// </summary>
        /// <param name="text"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        internal static List<GlyphRect> GetBoundsOfEachGlyph(string text, OpenTypeFont font)
        {
            //TODO: Should be cached to each font somehow
            List<GlyphRect> rects = new List<GlyphRect>();

            var glyphMappings = font.CmapTable.GetPreferredSubtable().GetGlyphMappings();

            //Dictionary<ushort, BoundingRectangle> glyphDict = new Dictionary<ushort, BoundingRectangle>();

            for (int i = 0; i < text.Count(); i++)
            {
                var c = text[i];

                if ((c == '\n' || c == '\r'))
                {
                    continue; //CRLF is irrelevant for getting the glyph bounding boxes
                }

                var gi = glyphMappings.GetGlyphIndex(c);
                double gWidth;
                double advanceWidth = 0;

                var hhMetric = font.HmtxTable.hMetrics[gi ?? 0];

                if (gi == 0)
                {
                    advanceWidth = font.Os2Table.xAvgCharWidth;
                }
                else
                {
                    advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                }

                try
                {
                    var leftSideBearing = hhMetric.lsb;
                    if (leftSideBearing < 0)
                    {
                        //if side bearing negative the glyph takes more space than the advancewidth
                        //We are missing rsb or reference to glyf table
                        gWidth = advanceWidth - leftSideBearing;
                    }
                    else
                    {
                        gWidth = advanceWidth;
                    }

                    gWidth /= (double)font.HeadTable.UnitsPerEm;

                    //TODO: Get glyph height. For now. Assume EM-height
                    var gRect = new GlyphRect(gi.Value, gWidth, font.FullName);
                    rects.Add(gRect);
                }
                catch (Exception ex)
                {

                    throw new Exception(ex.Message);
                }
            }

            return rects;
        }

        internal static double MeasureAscent(OpenTypeFont font, double fontSize)
        {
            return font.Os2Table.usWinAscent * (fontSize / font.HeadTable.UnitsPerEm);
        }

        internal static double MeasureDescent(OpenTypeFont font, double fontSize)
        {
            //TypoDescender is Short WinDescent ushort so we use long so both can be assigned same variable
            long descentBase = font.Os2Table.UseTypoMetrics ? -(long)font.Os2Table.sTypoDescender : (long)font.Os2Table.usWinDescent;

            return descentBase * (fontSize / font.HeadTable.UnitsPerEm);
        }

        /// <summary>
        /// Returns advanceWidth for char
        /// </summary>
        /// <param name="glyphMappings"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        private static int CalcGlyphWidth(GlyphMappings glyphMappings, char c, OpenTypeFont fontData, ref ushort? lastGlyphIndex, ref bool applyKerning)
        {
            int advanceWidth;

            var gi = glyphMappings.GetGlyphIndex(c);

            if (gi == 0 && c != 0)
            {
                advanceWidth = fontData.Os2Table.xAvgCharWidth;
            }
            else
            {
                var hhMetric = fontData.HmtxTable.hMetrics[gi ?? 0];
                advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
            }

            if (applyKerning)
            {
                var kern = GetKerningAdjustment(lastGlyphIndex ?? 0, gi ?? 0, fontData);
                advanceWidth += kern;
            }
            applyKerning = true;

            lastGlyphIndex = gi;

            return advanceWidth;
        }

        private static void MeasureAndWrapLines(string text, ref int totalAdvanceWidth, ref int totalWordWidth, OpenTypeFont fontData, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double maxWidth, List<string> wrappedStrings, bool applyKerning = true)
        {
            //Handle line-endings
            var splitStrings = text.Split([Environment.NewLine], StringSplitOptions.None);

            if (splitStrings.Length != 0)
            {
                //Avoid using kerning for first char/line
                MeasureAndWrapLine(splitStrings[0], fontData, ref totalAdvanceWidth, ref totalWordWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings, false);

                if (splitStrings.Length > 1)
                {
                    for (int i = 1; i < splitStrings.Count(); i++)
                    {
                        MeasureAndWrapLine(splitStrings[i], fontData, ref totalAdvanceWidth, ref totalWordWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings);
                    }
                }
            }
        }

        private static int CalculateAdvanceWidth(char c, GlyphMappings mappings, OpenTypeFont font, ref ushort? lastGlyphIndex, ref int lineWidth, ref int wordWidth, ref bool applyKerning)
        {
            int advanceWidth = CalcGlyphWidth(mappings, c, font, ref lastGlyphIndex, ref applyKerning);
            lineWidth += advanceWidth;

            wordWidth = c == ' ' ? 0 : wordWidth + advanceWidth;

            return advanceWidth;
        }

        private static void WrapAtCharPos(string line, int charPos, ref int nextLineStartIndex, ref int lineWidth, ref int wordWidth, int advanceWidth, List<string> wrappedStrings)
        {
            var wrappedString = ExtractWrappedSubstring(line, charPos, ref nextLineStartIndex, out TotalAdvanceMode advanceMode);
            wrappedStrings.Add(wrappedString);

            //Using enum to make it one Input parameter in WrapString instead of all 3
            //this as they're not actually used in there
            lineWidth = GetAdvanceWidthFromMode(advanceWidth, wordWidth, advanceMode);
            //New line means both totals are equal
            wordWidth = lineWidth;
        }

        private static void MeasureAndWrapLine(string line, OpenTypeFont font, ref int lineWidth, ref int wordWidth, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double maxWidth, List<string> wrappedStrings, bool applyKerning = true)
        {
            int nextLineStartIndex = 0;
            
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                var advanceWidth = CalculateAdvanceWidth(c, glyphMappings, font, ref lastGlyphIndex, ref lineWidth, ref wordWidth, ref applyKerning);

                if (lineWidth >= maxWidth)
                {
                    WrapAtCharPos(line, i, ref nextLineStartIndex, ref lineWidth, ref wordWidth, advanceWidth, wrappedStrings);
                }
            }

            var remainingLine = line.Substring(nextLineStartIndex);
            wrappedStrings.Add(remainingLine);

            //Has to be done After instead of before for loop.
            //For the case that we enter with an existing line width
            lineWidth = 0;
        }

        /// <summary>
        /// Measures the text and breaks it into smaller strings so that none exceed the MaxWidth
        /// </summary>
        /// <param name="text"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontData"></param>
        /// <param name="maxWidth"></param>
        /// <param name="preExistingLineWidth">point size previously calculated width of starting line</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        internal static List<string> MeasureAndWrapText(string text, double fontSize, OpenTypeFont fontData, double maxWidth, double preExistingLineWidth = 0, double preExistingWordWidth = 0)
        {
            //  Initialize:
            int totalAdvanceWidth = 0;
            ushort? lastGlyphIndex = 0;

            //Initalise collection to return
            List<string> wrappedStrings = new List<string>();

            var inputMaxWidth = maxWidth;
            //Convert maxWidth from points to font design units
            var maxWidthInDesignUnits = Math.Round(((inputMaxWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize), 0, MidpointRounding.AwayFromZero);

            var glyphMappings = fontData.CmapTable.GetPreferredSubtable().GetGlyphMappings();

            int wordWidth = 0;

            //If the starting line width is not zero.
            //Happens e.g. if other chars on the starting line have been measured with a different font.
            if (preExistingLineWidth != 0)
            {
                //Convert from points to current fonts font design units
                totalAdvanceWidth = Convert.ToInt16((preExistingLineWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize);
            }
            if (preExistingWordWidth != 0)
            {
                wordWidth = Convert.ToInt16((preExistingWordWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize);
            }

            // Execute:

            MeasureAndWrapLines(text, ref totalAdvanceWidth, ref wordWidth, fontData, glyphMappings, lastGlyphIndex, maxWidthInDesignUnits, wrappedStrings);

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
        internal static double MeasureText(string text, double fontSize, OpenTypeFont fontData, bool wrapText = false)
        {
            double totalAdvanceWidth = 0;
            ushort? lastGlyphIndex = 0;
            bool firstChar = true;

            double largestWidth = 0;

            var glyphMappings = fontData.CmapTable.GetPreferredSubtable().GetGlyphMappings();
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                var gi = glyphMappings.GetGlyphIndex(c);
                int advanceWidth;
                if (gi == 0 && c != 0)
                {
                    advanceWidth = fontData.Os2Table.xAvgCharWidth;
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi ?? 0];
                    advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                }

                if ((c == '\n' || c == '\r'))
                {
                    if (i > 0 && c == '\r' && text[i - 1] == '\n')
                    {
                        continue; //CRLF should be handled
                                  //as one new line.
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
                    int kerning = GetKerningAdjustment(lastGlyphIndex ?? 0, gi ?? 0, fontData);
                    totalAdvanceWidth += kerning;
                }

                lastGlyphIndex = gi;
                firstChar = false;
            }

            largestWidth = largestWidth < totalAdvanceWidth ? totalAdvanceWidth : largestWidth;

            // Convert to points
            return (largestWidth / (double)fontData.HeadTable.UnitsPerEm) * fontSize;
        }

        private static int GetKerningAdjustment(ushort left, ushort right, OpenTypeFont fontData)
        {
            if (fontData.KernTable != null)
            {
                foreach (var subtable in fontData.KernTable.SubTables)
                {
                    if (subtable.Format0Subtable == null) continue;
                    // Format 0 only
                    int format = subtable.coverage.RawValue >> 8;
                    bool isHorizontal = (subtable.coverage.RawValue & 0x1) == 1;
                    if (format != 0 || !isHorizontal) continue;
                    KerningPair[] pairs = subtable.Format0Subtable.Pairs;
                    if (pairs == null) continue;

                    //Left is high-order/most significant
                    //but since big-endian we must do it like this.
                    var combined = ((uint)left << 16) | right;

                    var index = OptimizedBinarySearch(pairs, combined, pairs.Length);
                    if (index < 0)
                    {
                        index = ~index;
                    }

                    if (index >= 0)
                    {
                        var maxIndex = pairs.Count();
                        if (maxIndex > index)
                        {
                            var pairItem = pairs[index];

                            //Extra verification in case something has gone wrong/the exact item does not exist
                            if (pairItem.left == left && pairItem.right == right)
                            {
                                return pairItem.value;
                            }
                            //else
                            //{
                            //    foreach (var pair in pairs)
                            //    {
                            //        if (pair.right == right && pair.left == left)
                            //        {
                            //            return pair.value;
                            //        }
                            //    }
                            //}
                        }
                        else
                        {
                            //Index has gone beyond max index.
                            //This should never be possible unless the file has been read wrong or is corrupt...
                            //font Arial black appears corrupt or missing kerning table?
                            return 0;
                            throw new Exception("Impossible kerning table detected(!?)");
                        }
                    }
                }
            }
            return 0;
        }

        private static int OptimizedBinarySearch(KerningPair[] arr, uint targetCombined, int length)
        {
            if (length == 0) return -1;
            int low = 0, high = length - 1, mid;

            while (low <= high)
            {
                mid = (low + high) >> 1;

                if (targetCombined < arr[mid].Combined)
                    high = mid - 1;

                else if (targetCombined > arr[mid].Combined)
                    low = mid + 1;

                else
                    return mid;
            }
            return ~low;
        }

        enum TotalAdvanceMode
        {
            Zero,
            LatestCharOnly,
            FromLastWord,
        }

        /// <summary>
        /// Pick out and return a string at the current char index from an original string
        /// </summary>
        /// <param name="orgLine">Original line text</param>
        /// <param name="cIdx">current char index in original line</param>
        /// <param name="c">current character</param>
        /// <param name="startLineIdx">The starting char index of old and then the new line in orgLine</param>
        /// <param name="mode">Informs calling method what the advance of the next line should be set to</param>
        /// <returns></returns>
        private static string ExtractWrappedSubstring(string orgLine, int cIdx, ref int startLineIdx, out TotalAdvanceMode mode)
        {
            //Result string
            string wrappedString = string.Empty;

            var prevStartIdx = startLineIdx;
            char c = orgLine[cIdx];

            //Number of chars in the current line
            var charCountFromLast = cIdx - prevStartIdx;

            //Substring out the current line from the original
            var txt = orgLine.Substring(prevStartIdx, charCountFromLast);

            //Ensure whole words get moved down if part of its letters are overflowing
            var splitLines = txt.Split(' ');

            if (splitLines.Length > 1 && c != ' ')
            {
                var stringOverMax = splitLines.Last();
                var startIndex = txt.Length - stringOverMax.Length;
                //Remove the overflowing characters
                var spacedString = txt.Remove(startIndex, stringOverMax.Length).TrimEnd(' ');

                wrappedString = spacedString;

                //The start index of the first character in the overflow (After space)
                startLineIdx += startIndex;

                mode = TotalAdvanceMode.FromLastWord;
            }
            else
            {
                //If the char was a space it should not be added to the next line
                //Therefore we do not add its width and the index of the next line starts at the next character.
                if (c == ' ')
                {
                    //The current char has crossed the max
                    //Therefore remove it from the text to be added.
                    wrappedString = txt.Substring(0, txt.Length);
                    startLineIdx = cIdx + 1;
                    mode = TotalAdvanceMode.Zero;
                }
                else
                {
                    //The current char has crossed the max
                    //Therefore remove it from the text to be added.
                    wrappedString = txt.Substring(0, txt.Length);
                    //The current character is part of the new line
                    //We should start at the index of the current character and add its width to the new line
                    startLineIdx = cIdx;
                    mode = TotalAdvanceMode.LatestCharOnly;
                }
            }

            return wrappedString;
        }

        private static int GetAdvanceWidthFromMode(int widthChar, int widthWord, TotalAdvanceMode mode)
        {
            switch (mode)
            {
                case TotalAdvanceMode.Zero:
                    return 0;
                case TotalAdvanceMode.LatestCharOnly:
                    return widthChar;
                case TotalAdvanceMode.FromLastWord:
                    return widthWord;
            }

            throw new InvalidOperationException($"AdvanceMode '{mode}' is not a valid advance mode. " +
                $"And does not exist within the enum: '{typeof(TotalAdvanceMode)}' ");
        }

        /// <summary>
        /// Converts the design units of one font to the design units of another font
        /// </summary>
        /// <param name="origFont"></param>
        /// <param name="origSize"></param>
        /// <param name="targetFont"></param>
        /// <param name="targetSize"></param>
        /// <param name="maxWidth"></param>
        /// <param name="lineWidth"></param>
        /// <param name="wordWidth"></param>
        internal static void ConvertDesignUnits(OpenTypeFont origFont, double origSize, OpenTypeFont targetFont, double targetSize, ref double maxWidth, ref int lineWidth, ref int wordWidth)
        {
            //Potential future optimization: Check if units perEm are equal if they are (most fonts are)
            //Should be able to only apply a factor of origSize/targetSize
            var factorOrig = origSize / ((double)origFont.HeadTable.UnitsPerEm);

            var maxWidthInPoints = maxWidth * factorOrig;
            var lineWidthInPoints = lineWidth * factorOrig;
            var wordWidthInPoints = wordWidth * factorOrig;

            var factorTarget = ((double)targetFont.HeadTable.UnitsPerEm) / targetSize;

            maxWidth = maxWidthInPoints * factorTarget;
            lineWidth = Convert.ToInt16(lineWidthInPoints * factorTarget);
            wordWidth = Convert.ToInt16(wordWidthInPoints * factorTarget);
        }

        private static void LogFragmentWidth(TextFragmentCollection collection, int UPM, double fontSize, int lineWidth, int prevFragmentWidths, int fragmentIdx)
        {
            var currentFragmentLineWidth = lineWidth - prevFragmentWidths;
            var factorOrig = fontSize / ((double)UPM);
            var lineWidthInPoints = currentFragmentLineWidth * factorOrig;

            collection.AddFragmentWidth(fragmentIdx, lineWidthInPoints);
        }

        private static void LogLineData(TextParagraph paragraph, OpenTypeFont currentFont, OpenTypeFont largestFont, double CurrentLineLargestSize, int lineWidth, int fragmentIdx)
        {
            //Log line width
            paragraph.Fragments.AddLineWidth(lineWidth / currentFont.HeadTable.UnitsPerEm * paragraph.FontSizes[fragmentIdx]);

            //Log line height
            paragraph.Fragments.AddLargestFontSizePerLine(CurrentLineLargestSize);
            var lineAscent = GetBaseLine(largestFont, CurrentLineLargestSize);
            paragraph.Fragments.AddAscentPerLine(lineAscent);
            var lineDescent = MeasureDescent(largestFont, CurrentLineLargestSize);
            paragraph.Fragments.AddDescentPerLine(lineDescent);
        }

        internal static List<string> WrapMultipleTextFragments(TextParagraph paragraph, double maxWidthPoints)
        {
            //Initialize variables
            List<string> wrappedStrings = new List<string>();
            ushort? lastGlyphIndex = null;
            bool applyKerning = false;
            OpenTypeFont currentFont = null;

            int lineWidth = 0;
            int wordWidth = 0;
            int prevLineEndIndex = 0;

            string leftOverLine = "";

            double maxWidth = maxWidthPoints;
            double fontSize = 0;

            //In the AllText string.
            //Does not take new text-strings into account
            int currentLineIndex = 0;

            //---Logger variables---
            var fragmentCollection = paragraph.Fragments;
            var allText = fragmentCollection.AllText;
            var newLineIndicies = fragmentCollection.AllTextNewLineIndicies;
            double CurrentLineLargestFontSize = 0;
            OpenTypeFont largestFontCurrentLine = paragraph.FontIndexDict[0];

            var prevFragmentIdx = 0;
            var prevFragWidths = 0;
            ////Technically currentLineOfFragmentWidth
            //int fragmentWidth = 0;
            ////Technically lastLINEofFragmentWidth
            ////lastFragmentWidth could be a different fragment OR current fragment on a previous line
            ////Whichever the previous measured piece of text was
            //int lastFragmentWidth = 0;
            //---Logger Variables end---

            //Iterate through All fragments as one concatenated text string
            for (int i = 0; i < allText.Length; i++)
            {
                //Get char info stored previously in the textParagraph class
                var charInfo = fragmentCollection.CharLookup[i];
                var lineIdx = charInfo.Line;
                var fragmentIdx = charInfo.Fragment;

                //Happens at the end of a fragment/start of a new fragment
                if(fragmentIdx != prevFragmentIdx)
                {
                    LogFragmentWidth(paragraph.Fragments, currentFont.HeadTable.UnitsPerEm, fontSize, lineWidth, prevFragWidths, prevFragmentIdx);

                    //Since there is no gurantee any wrapping has occured
                    //But we are moving on to the next fragment we need to know how much of the
                    //linewidth is not part of the current fragment
                    //Since there could be multiple fragments before line break add to current prevfragwidth
                    prevFragWidths += (lineWidth - prevFragWidths);
                }

                //If we hit a pre-existing line break. Reset line and wordwidths
                var indexExists = newLineIndicies.Count() > currentLineIndex;
                if(indexExists)
                {
                    if (i >= newLineIndicies[currentLineIndex])
                    {
                        //Log current line data
                        LogLineData(paragraph, currentFont, largestFontCurrentLine, CurrentLineLargestFontSize, lineWidth, fragmentIdx);
                        //Log current fragment width
                        LogFragmentWidth(paragraph.Fragments, currentFont.HeadTable.UnitsPerEm, fontSize, lineWidth, prevFragWidths, fragmentIdx);
                        //We are on the same fragment but there's been a line break
                        //Therefore the previous fragment width is 0 as it is within this fragment
                        prevFragWidths = 0;

                        //Since we are on a new line the current largest font-size for this line is the current size
                        CurrentLineLargestFontSize = paragraph.FontSizes[fragmentIdx];
                        largestFontCurrentLine = currentFont;

                        var addedLine = leftOverLine.Trim(['\r', '\n']);
                        wrappedStrings.Add(addedLine);
                        lineWidth = 0;
                        wordWidth = 0;
                        leftOverLine = "";
                        //prevLineEndIndex = i;
                        currentLineIndex++;
                    }
                }

                //If this char has a different font, do the neccesary conversions
                if (currentFont != null)
                {
                    if (currentFont != paragraph.FontIndexDict[fragmentIdx] | fontSize != paragraph.FontSizes[fragmentIdx])
                    {
                        ConvertDesignUnits(currentFont, fontSize, paragraph.FontIndexDict[fragmentIdx], paragraph.FontSizes[fragmentIdx],
                            ref maxWidth, ref lineWidth, ref wordWidth);
                        //Font/fragment change
                        currentFont = paragraph.FontIndexDict[fragmentIdx];
                        if (paragraph.FontSizes[fragmentIdx] > CurrentLineLargestFontSize)
                        {
                            largestFontCurrentLine = currentFont;
                            CurrentLineLargestFontSize = paragraph.FontSizes[fragmentIdx];
                        }
                    }
                }
                else
                {
                    currentFont = paragraph.FontIndexDict[fragmentIdx];
                    maxWidth = (maxWidthPoints * (double)currentFont.HeadTable.UnitsPerEm) / paragraph.FontSizes[fragmentIdx];
                    if (paragraph.FontSizes[fragmentIdx] > CurrentLineLargestFontSize)
                    {
                        largestFontCurrentLine = currentFont;
                        CurrentLineLargestFontSize = paragraph.FontSizes[fragmentIdx];
                    }
                }

                fontSize = paragraph.FontSizes[fragmentIdx];
                var glyphMapping = paragraph.GlyphMappings[fragmentIdx];

                char c = allText[i];
                var advanceWidth = CalculateAdvanceWidth(c, glyphMapping, currentFont, ref lastGlyphIndex, ref lineWidth, ref wordWidth, ref applyKerning);

                //Perform the actual wrapping
                if (lineWidth > maxWidth)
                {
                    WrapAtCharPos(allText, i, ref prevLineEndIndex, ref lineWidth, ref wordWidth, advanceWidth, wrappedStrings);

                    //Log where wrapping occured in order to keep track of fragment/run/richtext

                    //Can't be part of Log function as there could be pre-existing line-breaks
                    paragraph.Fragments.AddWrappingIndex(prevLineEndIndex);
                    //Log line data
                    LogLineData(paragraph, currentFont, largestFontCurrentLine, CurrentLineLargestFontSize, lineWidth, fragmentIdx);
                    //Log Fragment widths
                    LogFragmentWidth(paragraph.Fragments, currentFont.HeadTable.UnitsPerEm, fontSize, lineWidth, prevFragWidths, fragmentIdx);
                    prevFragWidths = 0;

                    //Since we're using the AllText, need to handle leftover line differently
                    if (i < prevLineEndIndex)
                    {
                        if(lineWidth != 0)
                        {
                            leftOverLine = allText.Substring(prevLineEndIndex, prevLineEndIndex - i);
                            //Since we've moved one beyond the last
                            i = prevLineEndIndex;
                        }
                        else
                        {
                            leftOverLine = "";
                        }
                    }
                    else
                    {
                        //Special case for only 1 or 0 chars in leftover line
                        leftOverLine = allText.Substring(prevLineEndIndex, i - prevLineEndIndex);
                        leftOverLine += allText.Substring(i, 1);
                    }
                    //Since we are on a new line the current largest font-size for this line is the current size
                    CurrentLineLargestFontSize = paragraph.FontSizes[fragmentIdx];
                    largestFontCurrentLine = currentFont;
                }
                else
                {
                    //Add the current char to current unwrapped line
                    leftOverLine += allText.Substring(i, 1);
                }

                prevFragmentIdx = fragmentIdx;
            }

            LogLineData(paragraph, currentFont, largestFontCurrentLine, CurrentLineLargestFontSize, lineWidth, paragraph.Fragments.TextFragments.Count()-1);
            LogFragmentWidth(paragraph.Fragments, currentFont.HeadTable.UnitsPerEm, fontSize, lineWidth, prevFragWidths, paragraph.Fragments.TextFragments.Count() - 1);

            wrappedStrings.Add(leftOverLine);

            return wrappedStrings;
        }
    }
}

namespace EPPlus.Fonts.OpenType
{
    public enum FontSubFamily
    {
        Regular,
        Bold,
        Italic,
        BoldItalic
    }
}