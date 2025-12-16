using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;


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

        private static void MeasureAndWrapLine(string line, OpenTypeFont fontData, int lineWidth, GlyphMappings glyphMappings, ushort? lastGlyphIndex, double maxWidth, List<string> wrappedStrings, bool applyKerning = true)
        {
            int nextLineStartIndex = 0;
            int wordWidth = 0;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                int advanceWidth = CalcGlyphWidth(glyphMappings, c, fontData, lastGlyphIndex, ref applyKerning);
                lineWidth += advanceWidth;

                //Diff starts
                wordWidth = c == ' ' ? 0 : wordWidth + advanceWidth;

                if (lineWidth > maxWidth)
                {
                    var wrappedString = WrapString(line, i, ref nextLineStartIndex, out TotalAdvanceMode advanceMode);
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

            //Has to be done after in case we enter with an existing line width
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

        private static double MeasureLine(string line, OpenTypeFont fontData, GlyphMappings glyphMappings, ushort? lastGlyphIndex, bool applyKerning = true)
        {
            int lineWidth = 0;

            for (int i = 0; i < line.Length; i++)
            {
                lineWidth += CalcGlyphWidth(glyphMappings, line[i], fontData, lastGlyphIndex, ref applyKerning);
            }

            return lineWidth;
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
        internal static List<string> MeasureAndWrapText(string text, double fontSize, OpenTypeFont fontData, double maxWidth, double preExistingLineWidth = 0)
        {
            int totalAdvanceWidth = 0;
            ushort? lastGlyphIndex = 0;
            var splitStrings = text.Split([Environment.NewLine], StringSplitOptions.None);

            //Initalise collection to return
            List<string> wrappedStrings = new List<string>();

            var inputMaxWidth = maxWidth;
            //Convert maxWidth from points to font design units
            maxWidth = (maxWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize;

            var glyphMappings = fontData.CmapTable.GetPreferredSubtable().GetGlyphMappings();

            //If the starting line width is not zero.
            //Happens e.g. if other chars on the starting line have been measured with a different font.
            if (preExistingLineWidth != 0)
            {
                totalAdvanceWidth = Convert.ToInt16((preExistingLineWidth * (double)fontData.HeadTable.UnitsPerEm) / fontSize);
            }

            //Avoid using kerning for first char/line
            MeasureAndWrapLine(splitStrings[0], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings, false);

            for (int i = 1; i < splitStrings.Count(); i++)
            {
                MeasureAndWrapLine(splitStrings[i], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings);
            }

            ////Avoid using kerning for first char/line
            //MeasureLine(splitStrings[0], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, false);

            //for (int i = 1; i < splitStrings.Count(); i++)
            //{
            //    MeasureLine(splitStrings[i], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth);
            //}

            ////Avoid using kerning for first char/line
            //MeasureAndWrapLine(splitStrings[0], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings, false);

            //for (int i = 1; i < splitStrings.Count(); i++)
            //{
            //    MeasureAndWrapLine(splitStrings[i], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings);
            //}
            //if (splitStrings.Length > 0 && splitStrings[0].Length > 0)
            //{
            //    //First glyph has no kerning
            //    var firstGlyphWidth = CalculateGlyphAdvanceWidth(glyphMappings, splitStrings[0][0], fontData, lastGlyphIndex, false);
            //    totalAdvanceWidth += firstGlyphWidth;

            //    for (int i = 1; i < splitStrings.Count(); i++)
            //    {
            //        var line = splitStrings[i];

            //        for (int j = 1; j < line[j]; j++)
            //        {
            //            char c = line[j];

            //            int advanceWidth = CalculateGlyphAdvanceWidth(glyphMappings, c, fontData, lastGlyphIndex);



            //            lineWidth += advanceWidth;

            //            if (MaxWidth != null && lineWidth > MaxWidth)
            //            {
            //                return lineWidth;
            //            }
            //        }

            //        lineWidth = 0;
            //    }
            //    ////Avoid using kerning for first char/line
            //    //MeasureAndWrapLine(splitStrings[0], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings, false);

            //    //for (int i = 1; i < splitStrings.Count(); i++)
            //    //{
            //    //    MeasureAndWrapLine(splitStrings[i], fontData, totalAdvanceWidth, glyphMappings, lastGlyphIndex, maxWidth, wrappedStrings);
            //    //}
            //}

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
        /// <param name="mode">Informs calling method what TotalAdvance should be set to</param>
        /// <returns></returns>
        private static string WrapString(string orgLine, int cIdx, ref int startLineIdx, out TotalAdvanceMode mode)
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

            //Split strings on line endings
            var newLine = Environment.NewLine;
            //Initalise collection to return
            List<string> wrappedStrings = new List<string>();

            //leftOverLine refers to a line of text that has not yet been wrapped
            string leftOverLine = "";
            double leftOverAdvanceWidthInPoints = 0;
            double leftOverTotalAdvanceFromLastWord = 0;

            var inputMaxWidth = maxWidth;

            var combinedString = string.Join(string.Empty, textFragments.ToArray());
            int charCount = 0;

            var splitStrings = combinedString.Split([newLine], StringSplitOptions.None);
            List<int> presetNewLineIndicies = new();
            int totalPresetLength = 0;

            foreach (var line in splitStrings)
            {
                presetNewLineIndicies.Add(totalPresetLength + line.Length);
                totalPresetLength += line.Length;
            }
            var currentPresetLineIndex = presetNewLineIndicies[0];

            int totalAdvanceFromLastWord = 0;

            for (int k = 0; k < textFragments.Count(); k++)
            {
                var testArray = splitStrings.Last().ToCharArray();

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
                    if (charCount >= currentPresetLineIndex)
                    {
                        totalAdvanceFromLastWord = 0;
                        totalAdvanceWidth = 0;
                        leftOverLine = "";
                        leftOverAdvanceWidthInPoints = 0;
                        leftOverTotalAdvanceFromLastWord = 0;
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

                    //We are beyond MaxWidth. Wrap the line.
                    if (newWidth > maxWidth)
                    {
                        var wrappedString = WrapString(combinedString, charCount, ref nextLineStartIndex, out TotalAdvanceMode advanceMode);

                        wrappedStrings.Add(wrappedString);

                        //Using enum to make it one Input parameter in WrapString instead of all 3
                        //this as they're not actually used in there
                        totalAdvanceWidth = GetAdvanceWidthFromMode(advanceWidth, totalAdvanceFromLastWord, advanceMode);
                        //New line means both totals are equal
                        totalAdvanceFromLastWord = totalAdvanceWidth;

                        //Since we're using the combined total string, need to handle leftover line differently
                        if (advanceMode != TotalAdvanceMode.FromLastWord)
                        {
                            leftOverLine = combinedString.Substring(nextLineStartIndex, nextLineStartIndex - charCount);
                        }
                        else
                        {
                            //Special case for only 1 or 0 chars in leftover line

                            //In this case CharCount will be equal to or larger than NextLineStartIndex (this as we skip ' ' and line endings)
                            //Meaning the "FromLastWord" substring would become impossible.
                            //Therefore we flip the minus it so that our leftover is either 1 or 0
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