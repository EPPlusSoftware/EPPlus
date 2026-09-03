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
  09/01/2026         EPPlus Software AB       Use shared cache and compact character lookup
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.GenericFontWidths;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal abstract class GenericFontMetricsTextMeasurerBase
    {
        /// <summary>
        /// Shared rather than per instance. The table holds 116 entries and never changes, and
        /// every measurer was allocating its own copy.
        /// </summary>
        private static readonly FontScaleFactors _fontScaleFactors = new FontScaleFactors();

        public GenericFontMetricsTextMeasurerBase()
        {
            // Nothing to do. Metrics load on first use through GenericFontMetricsCache, which
            // replaces the static dictionary this class used to populate eagerly - that read
            // every font in the archive the first time anything was measured.
        }

        internal protected bool IsValidFont(uint fontKey)
        {
            return GenericFontMetricsCache.IsValidFont(fontKey);
        }

        private static SerializedFontMetrics GetFont(uint fontKey)
        {
            var font = GenericFontMetricsCache.GetMetrics(fontKey);
            if (font == null)
            {
                throw new InvalidOperationException(
                    string.Format("No font metrics loaded for key {0}.", fontKey));
            }
            return font;
        }

        internal protected TextMeasurement MeasureTextInternal(string text, uint fontKey, MeasurementFontStyles style, float size, eWrappedTextAutofitMode mode = eWrappedTextAutofitMode.Skip)
        {
            if (text == null)
            {
                return new TextMeasurement(0, 0);
            }
            var sFont = GetFont(fontKey);

            // Width of the current segment (a "segment" is a line in SplitNewLine mode,
            // or a word in SplitWord mode). In FullText/Skip the whole text is one segment.
            var width = 0f;
            var widthEA = 0f;

            // Width of the widest segment seen so far.
            var maxWidth = 0f;
            var maxWidthEA = 0f;

            for (var x = 0; x < text.Length; x++)
            {
                var c = text[x];

                if (IsSegmentBoundary(c, mode))
                {
                    // A CRLF pair is a single line break, not two.
                    if (x > 0 && c == '\r' && text[x - 1] == '\n')
                    {
                        continue;
                    }

                    // A visible boundary character (hyphen) remains at the end of the
                    // segment it terminates, so its own width is added before the break.
                    FontMetricsClass boundaryClass;
                    if (IsVisibleBoundary(c) && sFont.TryGetClass(c, out boundaryClass))
                    {
                        width += sFont.GetClassWidth(boundaryClass);
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
                    continue;
                }

                //If east Asian char use default regardless of actual font.
                if (IsEastAsianChar(c))
                {
                    widthEA += GetEastAsianCharWidth(c, style);
                }
                else
                {
                    FontMetricsClass cls;
                    if (sFont.TryGetClass(c, out cls))
                    {
                        var fw = sFont.GetClassWidth(cls);
                        if (Char.IsDigit(c)) fw *= FontScaleFactors.DigitsScalingFactor;
                        width += fw;
                    }
                    else if (char.IsControl(c) == false)
                    {
                        width += sFont.DefaultWidth;
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

        internal List<uint> MeasureTextSpacingInternal(string text, uint fontKey, MeasurementFontStyles style, float size, float ppi = 108.73578912433f)
        {
            var sFont = GetFont(fontKey);
            var spacingBuffer = new List<uint>(text.Length);

            float resolutionDifference = ppi / 96f;
            float ptSize = size * (72f / 96f);
            float finalFactor = resolutionDifference * ptSize;

            for (var x = 0; x < text.Length; x++)
            {
                spacingBuffer.Add(MeasureCharacterInternal(sFont, text[x], ptSize, finalFactor));
            }

            return spacingBuffer;
        }

        internal uint MeasureCharacter(char c, uint fontKey, MeasurementFontStyles style, float ptSize, float finalFactor)
        {
            return MeasureCharacterInternal(GetFont(fontKey), c, ptSize, finalFactor);
        }

        private static uint MeasureCharacterInternal(SerializedFontMetrics sFont, char c, float ptSize, float finalFactor)
        {
            FontMetricsClass fntClass;
            if (!sFont.TryGetClass(c, out fntClass))
            {
                fntClass = sFont.DefaultWidthClass;
            }

            float adjustmentFactor = 0.012f * ptSize * ((int)fntClass);
            float deviceUnits = sFont.GetClassWidth(fntClass) * finalFactor - adjustmentFactor;

            return (uint)Math.Round(deviceUnits, MidpointRounding.AwayFromZero);
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

        /// <summary>
        /// Returns true when the character falls inside one of the Japanese/Kanji ranges.
        ///
        /// The LINQ version this replaces allocated an enumerator and a closure for every
        /// character measured. The early return covers Latin and most of the BMP below the CJK
        /// blocks, which is the overwhelming majority of cell content.
        /// </summary>
        private static bool IsEastAsianChar(char c)
        {
            var cc = (int)c;
            if (cc < 0x2E80) return false;

            foreach (var range in UniCodeRange.JapaneseKanji)
            {
                if (range.IsInRange(cc)) return true;
            }
            return false;
        }
    }
}