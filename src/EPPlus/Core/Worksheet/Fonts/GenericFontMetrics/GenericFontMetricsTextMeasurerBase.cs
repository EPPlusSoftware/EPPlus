/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.GenericFontWidths;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal abstract class GenericFontMetricsTextMeasurerBase
    {
        private FontScaleFactors _fontScaleFactors = new FontScaleFactors();
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static object _syncRoot = new object();

        public GenericFontMetricsTextMeasurerBase()
        {
            Initialize();
        }

        private static void Initialize()
        {
            lock (_syncRoot)
            {
                if (_fonts == null)
                {
                    _fonts = GenericFontMetricsLoader.LoadFontMetrics();
                }
            }
        }

        internal protected bool IsValidFont(uint fontKey)
        {
            return _fonts.ContainsKey(fontKey);
        }

        internal protected TextMeasurement MeasureTextInternal(string text, uint fontKey, MeasurementFontStyles style, float size, eWrappedTextAutofitMode mode = eWrappedTextAutofitMode.Skip)
        {
            if(text==null)
            {
               return new TextMeasurement(0, 0);
            }
            var sFont = _fonts[fontKey];

            // Width of the current segment (a "segment" is a line in SplitNewLine mode,
            // or a word in SplitWord mode). In FullText/Skip the whole text is one segment.
            var width = 0f;
            var widthEA = 0f;

            // Width of the widest segment seen so far.
            var maxWidth = 0f;
            var maxWidthEA = 0f;

            for (var x = 0; x < text.Length; x++)
            {
                var fnt = sFont;
                var c = text[x];

                if (IsSegmentBoundary(c, mode))
                {
                    // A CRLF pair is a single line break, not two.
                    if (x > 0 && c == '\r' && text[x - 1] == '\n')
                    {
                        continue; //CRLF should be handle
                                  //d as one new line.
                    }

                    // A visible boundary character (hyphen) remains at the end of the
                    // segment it terminates, so its own width is added before the break.
                    if (IsVisibleBoundary(c) && sFont.CharMetrics.ContainsKey(c))
                    {
                        width += fnt.ClassWidths[sFont.CharMetrics[c]];
                    }

                    // Close the current segment: keep it if it is the widest so far.
                    if ((width + widthEA) > (maxWidth + maxWidthEA))
                    {
                        maxWidth = width;
                        maxWidthEA = widthEA;
                    }

                    // Start a new, empty segment.
                    width = 0f;
                    widthEA = 0f;

                    // The boundary character itself is not part of the next segment.
                    // (Visible boundaries were already counted into the closed segment above.)
                    continue;
                }

                //If east Asian char use default regardless of actual font.
                if (IsEastAsianChar(c))
                {
                    widthEA += GetEastAsianCharWidth(c, style);
                }
                else
                {
                    if (sFont.CharMetrics.ContainsKey(c))
                    {
                        var fw = fnt.ClassWidths[sFont.CharMetrics[c]];
                        if (Char.IsDigit(c)) fw *= FontScaleFactors.DigitsScalingFactor;
                        width += fw;
                    }
                    else if (char.IsControl(c) == false)
                    {
                        width += sFont.ClassWidths[fnt.DefaultWidthClass];
                    }
                }
            }

            // Close the final segment.
            if ((width + widthEA) > (maxWidth + maxWidthEA))
            {
                maxWidth = width;
                maxWidthEA = widthEA;
            }

            width = maxWidth;
            widthEA = maxWidthEA;

            width *= size;
            widthEA *= size;
            var sf = _fontScaleFactors.GetScaleFactor(fontKey, width);
            width *= sf;
            width += widthEA;
            var height = sFont.LineHeight1em * size;
            return new TextMeasurement(width, height);
        }

        /// <summary>
        /// Returns true if the character ends the current measurement segment for the given mode.
        /// </summary>
        private static bool IsSegmentBoundary(char c, eWrappedTextAutofitMode mode)
        {
            switch (mode)
            {
                case eWrappedTextAutofitMode.SplitNewLine:
                    return c == '\n' || c == '\r';
                case eWrappedTextAutofitMode.SplitWord:
                    return c == '\n' || c == '\r' || c == ' ' || c == '\t'
                        || c == '\u002D'  // hyphen-minus
                        || c == '\u2010'; // hyphen
                default:
                    // FullText and Skip: the whole string is a single segment.
                    return false;
            }
        }

        /// <summary>
        /// Returns true if the boundary character is visible and therefore contributes
        /// its own width to the segment it terminates (hyphens). Whitespace and line
        /// breaks are invisible and contribute no width.
        /// </summary>
        private static bool IsVisibleBoundary(char c)
        {
            return c == '\u002D' || c == '\u2010';
        }

        static Dictionary<char, uint> AlphabetChars = new Dictionary<char, uint>
        {
            {'a', 0x06 },
            {'b', 0x07 },
            {'c', 0x05 },
            {'d', 0x07 },
            {'e', 0x06 },
            {'f', 0x04 },
            {'g', 0x07 },
            {'h', 0x07 },
            {'i', 0x03 },
            {'j', 0x03 },
            {'k', 0x06 },
            {'l', 0x03 },
            {'m', 0x09 },
            {'n', 0x07 },
            {'o', 0x07 },
            {'p', 0x07 },
            {'q', 0x07 },
            {'r', 0x04 },
            {'s', 0x05 },
            {'t', 0x04 },
            {'u', 0x07 },
            {'v', 0x05 },
            {'w', 0x09 },
            {'x', 0x05 },
            {'y', 0x05 },
            {'z', 0x05 },
            {'A', 0x07 },
            {'B', 0x06 },
            {'C', 0x07 },
            {'D', 0x08 },
            {'E', 0x06 },
            {'F', 0x06 },
            {'G', 0x08 },
            {'H', 0x08 },
            {'I', 0x03 },
            {'J', 0x04 },
            {'K', 0x06 },
            {'L', 0x05 },
            {'M', 0x0A },
            {'N', 0x08 },
            {'O', 0x09 },
            {'P', 0x06 },
            {'Q', 0x08 },
            {'R', 0x07 },
            {'S', 0x06 },
            {'T', 0x06 },
            {'U', 0x08 },
            {'V', 0x07 },
            {'W', 0x0B },
            {'X', 0x06 },
            {'Y', 0x05 },
            {'Z', 0x06 }
        };

        internal List<uint> MeasureTextSpacingInternal(string text, uint fontKey, MeasurementFontStyles style, float size, float ppi = 108.73578912433f)
        {
            var sFont = _fonts[fontKey];
            var chars = text.ToCharArray();

            var spacingBuffer = new List<uint>();

            var widthDefault = sFont.ClassWidths[sFont.DefaultWidthClass];

            float resolutionDifference = ppi / 96f;
            float ptSize = size * (72f / 96f);

            float finalFactor = resolutionDifference * ptSize;

            for (var x = 0; x < chars.Length; x++)
            {
                var fnt = sFont;
                var c = chars[x];

                var fntClass = sFont.CharMetrics.ContainsKey(c) ? sFont.CharMetrics[c] : fnt.DefaultWidthClass;
                float adjustmentFactor = 0.012f * ptSize * ((int)fntClass);

                float deviceUnits = fnt.ClassWidths[fntClass] * finalFactor - adjustmentFactor;

                var rounded = Math.Round(deviceUnits * 10, MidpointRounding.AwayFromZero);
                var final = rounded / 10;

                uint simplifiedWidth = (uint)(Math.Round(deviceUnits, MidpointRounding.AwayFromZero));
                spacingBuffer.Add(simplifiedWidth);
            }

            return spacingBuffer;
        }

        internal uint MeasureCharacter(char c, uint fontKey, MeasurementFontStyles style, float ptSize, float finalFactor)
        {
            var sFont = _fonts[fontKey];
            var fnt = sFont;

            var fntClass = sFont.CharMetrics.ContainsKey(c) ? sFont.CharMetrics[c] : fnt.DefaultWidthClass;
            float adjustmentFactor = 0.012f * ptSize * ((int)fntClass);

            float deviceUnits = fnt.ClassWidths[fntClass] * finalFactor - adjustmentFactor;

            uint simplifiedWidth = (uint)Math.Round(deviceUnits, MidpointRounding.AwayFromZero);
            return simplifiedWidth;
        }

        private static float GetEastAsianCharWidth(int cc, MeasurementFontStyles style)
        {
            var emWidth = (cc >= 65377 && cc <= 65439) ? 0.5f : 1f;
            if ((style & MeasurementFontStyles.Bold) != 0)
            {
                emWidth *= 1.05f;
            }
            return emWidth * (96F / 72F) * FontScaleFactors.JapaneseKanjiDefaultScalingFactor;
        }

        private static bool IsEastAsianChar(char c)
        {
            var cc = (int)c;

            return UniCodeRange.JapaneseKanji.Any(x => x.IsInRange(cc));
        }

    }
}
