using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using static EPPlus.Fonts.OpenType.TextData;
using static System.Net.Mime.MediaTypeNames;


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
            if(font.Os2Table.UseTypoMetrics)
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
                    //if (i > 0 && c == '\r' && text[i - 1] == '\n')
                    //{
                        continue; //CRLF is irrelevant for getting the glyph bounding boxes
                    //}
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
            return font.Os2Table.usWinDescent * (fontSize / font.HeadTable.UnitsPerEm);
        }

        /// <summary>
        /// Returns advanceWidth for char
        /// </summary>
        /// <param name="glyphMappings"></param>
        /// <param name="c"></param>
        /// <returns></returns>
        private static int CalcGlyphWidth(GlyphMappings glyphMappings, char c, OpenTypeFont fontData, ushort? lastGlyphIndex, ref bool applyKerning)
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
                advanceWidth += GetKerningAdjustment(lastGlyphIndex ?? 0, gi ?? 0, fontData);
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

                if(splitStrings.Length > 1)
                {
                    for (int i = 1; i < splitStrings.Count(); i++)
                    {
                        MeasureAndWrapLine(splitStrings[i], fontData, ref totalAdvanceWidth, ref totalWordWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings);
                    }
                }
            }
        }

        private static void MeasureAndWrapIndividualChar(string line, int charPos, ref int nextLineStartIndex, ref int lineWidth, ref int wordWidth, OpenTypeFont fontData, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double maxWidth, List<string> wrappedStrings, bool applyKerning = true)
        {
            var c = line[charPos];
            int advanceWidth = CalcGlyphWidth(glyphMappings, c, fontData, lastGlyphIndex, ref applyKerning);
            lineWidth += advanceWidth;

            wordWidth = c == ' ' ? 0 : wordWidth + advanceWidth;

            if (lineWidth > maxWidth)
            {
                var wrappedString = ExtractWrappedSubstring(line, charPos, ref nextLineStartIndex, out TotalAdvanceMode advanceMode);
                wrappedStrings.Add(wrappedString);

                //Using enum to make it one Input parameter in WrapString instead of all 3
                //this as they're not actually used in there
                lineWidth = GetAdvanceWidthFromMode(advanceWidth, wordWidth, advanceMode);
                //New line means both totals are equal
                wordWidth = lineWidth;
            }
        }

        private static void MeasureAndWrapLine(string line, OpenTypeFont fontData, ref int lineWidth, ref int wordWidth, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double maxWidth, List<string> wrappedStrings, bool applyKerning = true)
        {
            int nextLineStartIndex = 0;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                int advanceWidth = CalcGlyphWidth(glyphMappings, c, fontData, lastGlyphIndex, ref applyKerning);
                lineWidth += advanceWidth;

                wordWidth = c == ' ' ? 0 : wordWidth + advanceWidth;

                if (lineWidth > maxWidth)
                {
                    var wrappedString = ExtractWrappedSubstring(line, i, ref nextLineStartIndex, out TotalAdvanceMode advanceMode);
                    wrappedStrings.Add(wrappedString);

                    //Using enum to make it one Input parameter in WrapString instead of all 3
                    //this as they're not actually used in there
                    lineWidth = GetAdvanceWidthFromMode(advanceWidth, wordWidth, advanceMode);
                    //New line means both totals are equal
                    wordWidth = lineWidth;
                }
            }

            var remainingLine = line.Substring(nextLineStartIndex);
            wrappedStrings.Add(remainingLine);

            //Has to be done After instead of before for loop.
            //For the case that we enter with an existing line width
            lineWidth = 0;
        }

        //private static double[] MeasureLine(string line, OpenTypeFont fontData, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double MaxWidth = double.NaN, bool applyKerning = true)
        //{
        //    int lineWidth = 0;

        //    for (int i = 0; i < line.Length; i++)
        //    {
        //        char c = line[i];
        //        int advanceWidth = CalcGlyphWidth(glyphMappings, c, fontData, lastGlyphIndex, applyKerning);
        //        applyKerning = true;

        //        lineWidth += advanceWidth;

        //        if(MaxWidth != double.NaN && lineWidth > MaxWidth)
        //        {
        //            return new double[] {lineWidth, i};
        //        }
        //    }

        //    return new double[] {lineWidth};
        //}

        ///// <summary>
        ///// Measure line until max width
        ///// </summary>
        ///// <param name="line"></param>
        ///// <param name="fontData"></param>
        ///// <param name="glyphMappings"></param>
        ///// <param name="lastGlyphIndex"></param>
        ///// <param name="MaxWidth"></param>
        ///// <param name="applyKerning"></param>
        ///// <returns>
        ///// [0] lineWidth, (if arr.len == 1 it never passed max length)
        ///// [1] current advance width, 
        ///// [2] char index </returns>
        //private static double[] MeasureLine(string line, OpenTypeFont fontData, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double MaxWidth, bool applyKerning = true)
        //{
        //    int lineWidth = 0;

        //    for (int i = 0; i < line.Length; i++)
        //    {
        //        var advanceWidth = CalcGlyphWidth(glyphMappings, line[i], fontData, lastGlyphIndex, ref applyKerning);
        //        lineWidth += advanceWidth;
        //        if (MaxWidth != double.NaN && lineWidth > MaxWidth)
        //        {
        //            return new double[] { lineWidth, advanceWidth, i };
        //        }
        //    }

        //    return new double[] { lineWidth };
        //}

        private static double[] MeasureLine(string line, double fontSize, OpenTypeFont fontData, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double MaxWidth, ref double prevLineWidth, ref double prevWordWidth, bool applyKerning = true)
        {
            int lineWidth = 0;
            int wordWidth = 0;

            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];

                //If we are at a new line
                if ((c == '\n' || c == '\r'))
                {
                    //Cancel and return we are done with the current line
                    return new double[] { lineWidth, wordWidth, i };
                }

                var advanceWidth = CalcGlyphWidth(glyphMappings, c, fontData, lastGlyphIndex, ref applyKerning);
                lineWidth += advanceWidth;

                wordWidth = c == ' ' ? 0 : wordWidth + advanceWidth;

                if (MaxWidth != double.NaN && lineWidth > MaxWidth)
                {
                    return new double[] { lineWidth, wordWidth, i };
                }
            }

            return new double[] { lineWidth, wordWidth };
        }

        private static double MeasureLine(string line, OpenTypeFont fontData, double fontSize, GlyphMappings glyphMappings, ushort? lastGlyphIndex, bool applyKerning = true)
        {
            int lineWidth = 0;

            for (int i = 0; i < line.Length; i++)
            {
                lineWidth += CalcGlyphWidth(glyphMappings, line[i], fontData, lastGlyphIndex, ref applyKerning);
            }

            return (lineWidth / (double)fontData.HeadTable.UnitsPerEm) * fontSize;
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
            maxWidth = (maxWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize;

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

            MeasureAndWrapLines(text, ref totalAdvanceWidth, ref wordWidth, fontData, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings);

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

            ////For if we want to calculate the total glyph height within a specific string
            //short GreatestYMax = short.MinValue;
            //short LowestYMin = short.MaxValue;

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
                else
                {
                    ////First char has no kerning but it does have a left side value.
                    //var firstCharLsb = Convert.ToInt16(fontData.HmtxTable.hMetrics[gi].lsb);
                    //totalAdvanceWidth += firstCharLsb;
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

        private static List<int> GetStartIndicies(string stringsCombined)
        {
            List<int> combinedStartIndicies = new List<int>();

            var strings = stringsCombined.Split([Environment.NewLine], StringSplitOptions.None);
            var totalLength = 0;

            for (int i = 0; i < strings.Count(); i++)
            {
                combinedStartIndicies.Add(strings[i].Length + totalLength);
                totalLength += strings[i].Length;
            }

            return combinedStartIndicies;
        }

        internal static List<string> WrapMultipleTextFragments2(List<string> textFragments, List<double> fontSizes, Dictionary<double, OpenTypeFont> fonts, double maxWidth)
        {
            var combinedString = string.Join(string.Empty, textFragments.ToArray());

            var charCount = 0;
            List<char> currentLine = new List<char>();
            ushort? lastGlyphIndex = null;

            double currLineWidth = 0;
            double wordWidth = 0;
            bool applyKerning = false;

            for (int i = 0; i < textFragments.Count; i++)
            {
                var currentFont = fonts[i];
                var currentSize = fontSizes[i];
                var textFragment = textFragments[i];
                var glyphMappings = currentFont.CmapTable.GetPreferredSubtable().GetGlyphMappings();


                var lineData = MeasureLine(textFragment, currentSize, currentFont, glyphMappings, lastGlyphIndex, maxWidth, 
                    ref currLineWidth, ref wordWidth, applyKerning);

                //We've hit a newline or gone over maxwidth
                if (lineData.Length > 2)
                {
                    var isNewLine = lineData[0] < maxWidth;
    
                    if (isNewLine)
                    {
                        
                    }
                    else
                    {

                    }
                }
                else
                {

                }
                //var lineData = MeasureLine(textFragment, currentSize, currentFont, glyphMappings, lastGlyphIndex, maxWidth);

                //if(lineData.Length > 2)
                //{

                //}
                //else
                //{

                //}

                //charCount++;
                //for (int j = 0; j < textFragments[i].Length; j++)
                //{


                //    charCount++;
                //}
            }

            ////First, we need to be able to tell what fragment belongs where as we iterate
            //var combinedString = string.Join(string.Empty, textFragments.ToArray());
            //var startIndicies = GetStartIndicies(textFragments, combinedString);

            //var currentFragment = 0;
            ////Currently also 0
            //var currentStartCharIndex = startIndicies[currentFragment];
            //var nextStartCharIndex = startIndicies[currentFragment++];

            ////Then we wish to iterate, one char at a time
            //for (int i = 0; i < combinedString.Length; i++)
            //{
            //    if(i >= nextStartCharIndex)
            //    {
            //        currentFragment++;
            //        nextStartCharIndex++;
            //    }

            //    var c = combinedString[i];


            //}

            //for (int i = 0; i < combinedString.Length; i++)
            //{
            //    var c = combinedString[i];


            //}
            ////int charCount = 0;

            //var splitStrings = combinedString.Split([Environment.NewLine], StringSplitOptions.None);
            //List<int> presetNewLineIndicies = new();
            //int presetCharCount = 0;

            //foreach (var line in splitStrings)
            //{
            //    presetNewLineIndicies.Add(presetCharCount + line.Length);
            //    presetCharCount += line.Length;
            //}

            ////Initalise collection to return
            //List<string> wrappedStrings = new List<string>();

            //for (int i = 0; i < textFragments.Count; i++)
            //{

            //}

            //var combinedString = string.Join(string.Empty, textFragments.ToArray());
            //int charCount = 0;

            //for (int i = 0; i < textFragments.Count; i++)
            //{
            //    MeasureLine(textFragments[i], fonts[i], fonts[i].)
            //}
            //var splitStrings = combinedString.Split([Environment.NewLine], StringSplitOptions.None);

            //for (int k = 0; k < textFragments.Count(); k++)
            //{
            //    MeasureAndWrapText()   
            //}

            //var glyphMappings = fonts[0].CmapTable.GetPreferredSubtable().GetGlyphMappings();
            //var firstLines = MeasureAndWrapText(textFragments[0], fontSizes[0], fonts[0], maxWidth);

            //var lastLine = firstLines.Last();

            //var lastLineLength = MeasureLine(lastLine, fonts[0], fontSizes[0], glyphMappings, null, false);

            //for (int i = 0; i < firstLines.Count()-1; i++)
            //{
            //    wrappedStrings.Add(firstLines[i]);
            //}

            //List<string> currentLines = new List<string>();

            //var currentLine = wrappedStrings.Last();

            //for (int k = 0; k < textFragments.Count(); k++)
            //{
            //    glyphMappings = fonts[k].CmapTable.GetPreferredSubtable().GetGlyphMappings();

            //    currentLines = MeasureAndWrapText(textFragments[k], fontSizes[k], fonts[k], maxWidth, lastLineLength);

            //    for (int i = 0; i < currentLines.Count() - 1; i++)
            //    {
            //        wrappedStrings.Add(currentLines[i]);
            //    }

            //    lastLineLength = MeasureLine(firstLines.Last(), fonts[k], fontSizes[k], glyphMappings, null, false);
            //}

            //wrappedStrings.Add(currentLines.Last());

            //return wrappedStrings;

            throw new NotImplementedException("not yet done");
        }

        /// <summary>
        /// Wrap multiple text-fragments (such as text-runs) that contain more than one font/or font size and return the resulting lines.
        /// </summary>
        /// <param name="textFragments"></param>
        /// <param name="fontSizes"></param>
        /// <param name="fonts"></param>
        /// <param name="maxWidth"></param>
        /// <returns></returns>
        internal static List<string> WrapMultipleTextFragments(List<string> textFragments, List<double> fontSizes, Dictionary<double, OpenTypeFont> fonts, double maxWidth)
        {
            int totalAdvanceWidth = 0;
            ushort? lastGlyphIndex = 0;
            bool firstChar = true;

            //Initalise collection to return
            List<string> wrappedStrings = new List<string>();

            //leftOverLine refers to a line of text that has not yet been wrapped
            string leftOverLine = "";
            double leftOverAdvanceWidthInPoints = 0;
            double leftOverTotalAdvanceFromLastWord = 0;

            var inputMaxWidth = maxWidth;

            var combinedString = string.Join(string.Empty, textFragments.ToArray());
            int charCount = 0;

            var splitStrings = combinedString.Split([Environment.NewLine], StringSplitOptions.None);
            List<int> presetNewLineIndicies = new();
            int totalPresetLength = 0;

            foreach (var line in splitStrings)
            {
                presetNewLineIndicies.Add(totalPresetLength + line.Length);
                totalPresetLength += line.Length;
            }

            var currentPresetLineIndex = 0;
            int totalAdvanceFromLastWord = 0;

            for (int k = 0; k < textFragments.Count(); k++)
            {
                //Convert maxWidth from points to current font design units (different fonts can have different units)
                maxWidth = (inputMaxWidth * (double)fonts[k].HeadTable.UnitsPerEm) / fontSizes[k];

                var glyphMappings = fonts[k].CmapTable.GetPreferredSubtable().GetGlyphMappings();

                int nextLineStartIndex = 0;
                totalAdvanceFromLastWord = 0;

                if (leftOverAdvanceWidthInPoints != 0)
                {
                    //Convert leftOverWidth and widthFromLastWord to current font design units
                    totalAdvanceWidth = Convert.ToInt16((leftOverAdvanceWidthInPoints * (double)fonts[k].HeadTable.UnitsPerEm) / fontSizes[k]);
                    totalAdvanceFromLastWord = Convert.ToInt16((leftOverTotalAdvanceFromLastWord * (double)fonts[k].HeadTable.UnitsPerEm) / fontSizes[k]);
                }

                for (int i = 0; i < textFragments[k].Length; i++)
                {
                    //Text-Fragments may already contain new lines
                    //Reset all advance when we reach such a newLine
                    if (charCount >= presetNewLineIndicies[currentPresetLineIndex])
                    {
                        totalAdvanceFromLastWord = 0;
                        totalAdvanceWidth = 0;
                        leftOverLine = "";
                        leftOverAdvanceWidthInPoints = 0;
                        leftOverTotalAdvanceFromLastWord = 0;
                        currentPresetLineIndex = currentPresetLineIndex == presetNewLineIndicies.Count() - 1 ? currentPresetLineIndex : currentPresetLineIndex++;
                    }

                    char c = textFragments[k][i];
                    var gi = glyphMappings.GetGlyphIndex(c);
                    int advanceWidth;
                    if (gi == 0 && c != 0)
                    {
                        advanceWidth = fonts[k].Os2Table.xAvgCharWidth;
                    }
                    else
                    {
                        var hhMetric = fonts[k].HmtxTable.hMetrics[gi ?? 0];
                        advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                    }

                    var newWidth = totalAdvanceWidth + advanceWidth;

                    int kerning = 0;
                    // Kerning adjustment
                    if (!firstChar)
                    {
                        kerning = GetKerningAdjustment(lastGlyphIndex ?? 0, gi ?? 0, fonts[k]);
                        newWidth += kerning;
                    }

                    totalAdvanceFromLastWord += (advanceWidth + kerning);

                    if (c == ' ')
                    {
                        totalAdvanceFromLastWord = 0;
                    }

                    if (newWidth > maxWidth)
                    {
                        var wrappedString = ExtractWrappedSubstring(combinedString, charCount, ref nextLineStartIndex, out TotalAdvanceMode advanceMode);

                        wrappedStrings.Add(wrappedString);

                        //Using enum to make it one Input parameter in WrapString instead of all 3
                        //this as they're not actually used in there
                        totalAdvanceWidth = GetAdvanceWidthFromMode(advanceWidth, totalAdvanceFromLastWord, advanceMode);
                        //New line means both totals are equal
                        totalAdvanceFromLastWord = totalAdvanceWidth;

                        //Since we're using the combined total string, need to handle leftover line differently
                        if (charCount < nextLineStartIndex)
                        {
                            leftOverLine = combinedString.Substring(nextLineStartIndex, nextLineStartIndex - charCount);
                        }
                        else
                        {
                            //Special case for only 1 or 0 chars in leftover line
                            leftOverLine = combinedString.Substring(nextLineStartIndex, charCount - nextLineStartIndex);
                        }
                    }
                    else
                    {
                        totalAdvanceWidth = newWidth;
                    }

                    lastGlyphIndex = gi;
                    firstChar = false;

                    //Add the current char to current unwrapped line
                    leftOverLine += combinedString.Substring(charCount, 1);
                    charCount++;
                }

                //We are about to exit or enter a new text-fragment which may have a different font. Save current advance in points
                leftOverAdvanceWidthInPoints = (totalAdvanceWidth / (double)fonts[k].HeadTable.UnitsPerEm) * fontSizes[k];
                leftOverTotalAdvanceFromLastWord = (totalAdvanceFromLastWord / (double)fonts[k].HeadTable.UnitsPerEm) * fontSizes[k];
            }

            wrappedStrings.Add(leftOverLine);

            return wrappedStrings;
        }

        internal struct LineInfo()
        {
            internal int lineWidth = 0;
            internal int wordWidth = 0;
            //leftOverLine refers to a line of text that has not yet been wrapped
            internal string leftOverLine = "";
            internal double prevLineWidth = 0;
            internal double prevWordWidth = 0;
        }

        internal static void WrapFragments(TextParagraph paragraph, LineInfo lineInfo, double inputMaxWidth, double maxWidth, int charCount, int currentLineIndex, List<string> wrappedStrings)
        {
            bool applyKerning = false;
            ushort? lastGlyphIndex = 0;

            for (int i = 0; i < paragraph.TextFragments.Count(); i++)
            {
                //Get Font data from the stored data
                var font = paragraph.FontIndexDict[i];
                var glyphMappings = paragraph.GlyphMappings[i];
                var fontSize = paragraph.FontSizes[i];

                var currentFragment = paragraph.TextFragments[i];

                int nextLineStartIndex = 0;
                lineInfo.wordWidth = 0;

                //Alternative argument: lineInfo.prevLineWidth != 0
                if (paragraph.CharLookup[charCount].Line == currentLineIndex)
                {
                    //This Fragment and the last fragment are part of the same line
                    //Convert leftOverWidth and widthFromLastWord to current font design units
                    lineInfo.lineWidth = Convert.ToInt16((lineInfo.prevLineWidth * (double)font.HeadTable.UnitsPerEm) / fontSize);
                    lineInfo.wordWidth = Convert.ToInt16((lineInfo.prevWordWidth * (double)font.HeadTable.UnitsPerEm) / fontSize);
                }
                else
                {
                    //This Fragment starts on a new line from the previous fragment
                }

                //Convert maxWidth from points to current font design units (different fonts can have different units)
                maxWidth = (inputMaxWidth * (double)font.HeadTable.UnitsPerEm) / fontSize;

                HandleCharsInFragment(paragraph, i, currentFragment, charCount, currentLineIndex, lineInfo, wrappedStrings, nextLineStartIndex, lastGlyphIndex, maxWidth, applyKerning);

                //We are about to exit or enter a new text-fragment which may have a different font. Save current advance in points
                lineInfo.prevLineWidth = (lineInfo.lineWidth / (double)font.HeadTable.UnitsPerEm) * fontSize;
                lineInfo.prevWordWidth = (lineInfo.wordWidth / (double)font.HeadTable.UnitsPerEm) * fontSize;
            }
        }

        internal static void HandleCharsInFragment(TextParagraph paragraph, int fragmentIdx, string fragment, int charCount, int currentLineIndex, LineInfo lineInfo, List<string> wrappedStrings, int nextLineStartIndex, 
            ushort? lastGlyphIndex, double maxWidth, bool applyKerning)
        {
            for (int j = 0; j < fragment.Length; j++)
            {
                var charInfo = paragraph.CharLookup[charCount];

                //Text-Fragments may already contain new lines
                //Reset all advance when we reach such a newLine
                if (charInfo.Line > currentLineIndex)
                {
                    lineInfo = new LineInfo();
                    currentLineIndex = charInfo.Line;
                }

                var wrappedStringCount = wrappedStrings.Count;

                MeasureAndWrapIndividualChar(fragment, j, ref nextLineStartIndex, ref lineInfo.lineWidth, ref lineInfo.wordWidth, paragraph.FontIndexDict[fragmentIdx],
                     paragraph.GlyphMappings[fragmentIdx], lastGlyphIndex, maxWidth, wrappedStrings, applyKerning);

                if (wrappedStringCount < wrappedStrings.Count)
                {
                    //Since we're using the combined total string, need to handle leftover line differently
                    if (charCount < nextLineStartIndex)
                    {
                        lineInfo.leftOverLine = paragraph.AllText.Substring(nextLineStartIndex, nextLineStartIndex - charCount);
                    }
                    else
                    {
                        //Special case for only 1 or 0 chars in leftover line
                        lineInfo.leftOverLine = paragraph.AllText.Substring(nextLineStartIndex, charCount - nextLineStartIndex);
                    }
                }

                //Add the current char to current unwrapped line
                lineInfo.leftOverLine += paragraph.AllText.Substring(charCount, 1);
                charCount++;
            }
        }

        internal static List<string> WrapMultipleTextFragments4(TextParagraph paragraph, double maxWidth)
        {
            //Keep track of original maxwidth in points
            var inputMaxWidth = maxWidth;
            int currentLineIndex = 0;

            //Initalise collection to return
            List<string> wrappedStrings = new List<string>();

            //Create struct to hold onto info from the previous fragment/earlier data on current line
            var lineInfo = new LineInfo();

            int charCount = 0;

            WrapFragments(paragraph, lineInfo, inputMaxWidth, maxWidth, charCount, currentLineIndex, wrappedStrings);

            wrappedStrings.Add(lineInfo.leftOverLine);

            return wrappedStrings;
        }

        internal static void ConvertBetweenFonts(OpenTypeFont origFont, double origSize, OpenTypeFont targetFont, double targetSize, ref double maxWidth, ref int lineWidth, ref int wordWidth)
        {
            //Potential future optimization: Check if units perEm are equal if they are (most fonts are)
            //Should be able to only apply a factor of origSize/targetSize

            var factorOrig = (double)origFont.HeadTable.UnitsPerEm * origSize;

            var maxWidthInPoints = maxWidth / factorOrig;
            var lineWidthInPoints = lineWidth / factorOrig;
            var wordWidthInPoints = wordWidth / factorOrig;

            var factorTarget = (double)targetFont.HeadTable.UnitsPerEm / targetSize;

            maxWidth = Convert.ToInt16(maxWidthInPoints * factorTarget);
            lineWidth = Convert.ToInt16(lineWidthInPoints * factorTarget);
            wordWidth = Convert.ToInt16(wordWidthInPoints * factorTarget);
        }

        internal static List<string> WrapMultipleTextFragments5(TextParagraph paragraph, double maxWidthPoints)
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

            //Iterate through All fragments as one concatenated text string
            for(int i = 0; i < paragraph.AllText.Length; i++)
            {
                //Get char info stored previously in the textParagraph class
                var charInfo = paragraph.CharLookup[i];
                var lineIdx = charInfo.Line;
                var fragmentIdx = charInfo.Fragment;

                //If we hit a pre-existing line break. Reset line and wordwidths
                if (i >= paragraph.AllTextNewLineIndicies[currentLineIndex])
                {
                    lineWidth = 0;
                    wordWidth = 0;
                    leftOverLine = "";
                    //prevLineEndIndex = i;
                    currentLineIndex++;
                }

                //If this char has a different font, do the neccesary conversions
                if(currentFont != null)
                {
                    if (currentFont != paragraph.FontIndexDict[fragmentIdx] | fontSize != paragraph.FontSizes[fragmentIdx])
                    {
                        ConvertBetweenFonts(currentFont, fontSize, paragraph.FontIndexDict[fragmentIdx], paragraph.FontSizes[fragmentIdx], 
                            ref maxWidth, ref lineWidth, ref wordWidth);
                        //Font/fragment change
                        //TODO: change from fontDesign units linewidth to points and back into new context
                        currentFont = paragraph.FontIndexDict[fragmentIdx];
                    }
                }
                else
                {
                    currentFont = paragraph.FontIndexDict[fragmentIdx];
                    var factorTarget = (double)currentFont.HeadTable.UnitsPerEm / paragraph.FontSizes[fragmentIdx];
                    maxWidth = Convert.ToInt16(maxWidthPoints * factorTarget);
                }

                fontSize = paragraph.FontSizes[fragmentIdx];

                var c = paragraph.AllText[i];

                //Calculate the width of the char/glyph
                int advanceWidth = CalcGlyphWidth(paragraph.GlyphMappings[fragmentIdx], c, currentFont, lastGlyphIndex, ref applyKerning);

                //Update advance and word widths
                lineWidth += advanceWidth;
                wordWidth = c == ' ' ? 0 : wordWidth + advanceWidth;

                //Perform the actual wrapping
                if (lineWidth > maxWidth)
                {
                    var wrappedString = ExtractWrappedSubstring(paragraph.AllText, i, ref prevLineEndIndex, out TotalAdvanceMode advanceMode);
                    wrappedStrings.Add(wrappedString);

                    //Using enum to make it one Input parameter in WrapString instead of all 3
                    //this as they're not actually used in there
                    lineWidth = GetAdvanceWidthFromMode(advanceWidth, wordWidth, advanceMode);

                    //New line means both totals are equal
                    wordWidth = lineWidth;

                    //Since we're using the AllText, need to handle leftover line differently
                    if (i < prevLineEndIndex)
                    {
                        leftOverLine = paragraph.AllText.Substring(prevLineEndIndex, prevLineEndIndex - i);
                    }
                    else
                    {
                        //Special case for only 1 or 0 chars in leftover line
                        leftOverLine = paragraph.AllText.Substring(prevLineEndIndex, i - prevLineEndIndex);
                    }
                }
            }

            wrappedStrings.Add(leftOverLine);

            return wrappedStrings;
        }

        internal static List<string> WrapMultipleTextFragments3(List<string> textFragments, List<double> fontSizes, Dictionary<double, OpenTypeFont> fonts, double maxWidth)
        {
            //Keep track of original maxwidth in points
            var inputMaxWidth = maxWidth;

            var allText = string.Join(string.Empty, textFragments.ToArray());
            int charCount = 0;

            //Get the indicies where newlines occur in the combined string
            List<int> allTextNewLinesIndicies = GetStartIndicies(allText);

            //Index for what line we are currently at
            int currentPresetLineIndex = -1;

            //Create struct to hold onto info from the previous fragment/earlier data on current line
            var lineInfo = new LineInfo();
            currentPresetLineIndex++;

            ushort? lastGlyphIndex = 0;
            bool applyKerning = false;

            //Initalise collection to return
            List<string> wrappedStrings = new List<string>();

            for (int k = 0; k < textFragments.Count(); k++)
            {
                //Convert maxWidth from points to current font design units (different fonts can have different units)
                maxWidth = (inputMaxWidth * (double)fonts[k].HeadTable.UnitsPerEm) / fontSizes[k];
                var glyphMappings = fonts[k].CmapTable.GetPreferredSubtable().GetGlyphMappings();

                int nextLineStartIndex = 0;
                lineInfo.wordWidth = 0;

                if (lineInfo.prevLineWidth != 0)
                {
                    //Convert leftOverWidth and widthFromLastWord to current font design units
                    lineInfo.lineWidth = Convert.ToInt16((lineInfo.prevLineWidth * (double)fonts[k].HeadTable.UnitsPerEm) / fontSizes[k]);
                    lineInfo.wordWidth = Convert.ToInt16((lineInfo.prevWordWidth * (double)fonts[k].HeadTable.UnitsPerEm) / fontSizes[k]);
                }

                for (int i = 0; i < textFragments[k].Length; i++)
                {
                    ////Text-Fragments may already contain new lines
                    ////Reset all advance when we reach such a newLine
                    if (charCount >= allTextNewLinesIndicies[currentPresetLineIndex])
                    {
                        lineInfo = new LineInfo();
                        currentPresetLineIndex = currentPresetLineIndex == allTextNewLinesIndicies.Count() - 1 ? currentPresetLineIndex : currentPresetLineIndex++;
                    }

                    var wrappedStringCount = wrappedStrings.Count;

                    MeasureAndWrapIndividualChar(textFragments[k], i, ref nextLineStartIndex, ref lineInfo.lineWidth, ref lineInfo.wordWidth, fonts[k], 
                        glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings, applyKerning);
                    
                    if(wrappedStringCount < wrappedStrings.Count)
                    {
                        //Since we're using the combined total string, need to handle leftover line differently
                        if (charCount < nextLineStartIndex)
                        {
                            lineInfo.leftOverLine = allText.Substring(nextLineStartIndex, nextLineStartIndex - charCount);
                        }
                        else
                        {
                            //Special case for only 1 or 0 chars in leftover line
                            lineInfo.leftOverLine = allText.Substring(nextLineStartIndex, charCount - nextLineStartIndex);
                        }
                    }

                    //Add the current char to current unwrapped line
                    lineInfo.leftOverLine += allText.Substring(charCount, 1);
                    charCount++;
                }

                //We are about to exit or enter a new text-fragment which may have a different font. Save current advance in points
                lineInfo.prevLineWidth = (lineInfo.lineWidth / (double)fonts[k].HeadTable.UnitsPerEm) * fontSizes[k];
                lineInfo.prevWordWidth = (lineInfo.wordWidth / (double)fonts[k].HeadTable.UnitsPerEm) * fontSizes[k];
            }

            wrappedStrings.Add(lineInfo.leftOverLine);

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