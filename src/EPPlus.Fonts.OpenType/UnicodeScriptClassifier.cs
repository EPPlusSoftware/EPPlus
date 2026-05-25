/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2026         EPPlus Software AB           Initial implementation
  05/20/2026         EPPlus Software AB           Auto-sort ranges so source can stay grouped by script
  05/25/2026         EPPlus Software AB           Added CJK Symbols/Punctuation and Halfwidth/Fullwidth Forms
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Classifies Unicode code points into <see cref="UnicodeScript"/> values.
    /// Used during text shaping to route glyphs missing from the primary font to the
    /// appropriate per-script fallback chain.
    ///
    /// Lookup is performed by binary search over a sorted range table, giving O(log n)
    /// classification per code point. The table covers all supported scripts; code points
    /// outside any range return <see cref="UnicodeScript.Unknown"/>.
    ///
    /// The source-code form of the table is grouped by script for readability. The actual
    /// search table is built once at type initialization by sorting a copy by Start.
    /// </summary>
    internal static class UnicodeScriptClassifier
    {
        /// <summary>
        /// Returns the Unicode script that contains the given code point, or
        /// <see cref="UnicodeScript.Unknown"/> if the code point is not in any supported range.
        /// </summary>
        public static UnicodeScript OfCodePoint(uint codePoint)
        {
            // Binary search for the range containing codePoint. Ranges are sorted by Start
            // and are non-overlapping.
            int lo = 0;
            int hi = _sortedRanges.Length - 1;

            while (lo <= hi)
            {
                int mid = lo + ((hi - lo) >> 1);
                var range = _sortedRanges[mid];

                if (codePoint < range.Start)
                {
                    hi = mid - 1;
                }
                else if (codePoint > range.End)
                {
                    lo = mid + 1;
                }
                else
                {
                    return range.Script;
                }
            }

            return UnicodeScript.Unknown;
        }

        private readonly struct Range
        {
            public readonly uint Start;
            public readonly uint End;
            public readonly UnicodeScript Script;

            public Range(uint start, uint end, UnicodeScript script)
            {
                Start = start;
                End = end;
                Script = script;
            }
        }

        // Source-form table — grouped by script for readability. Order does NOT matter here;
        // the actual search table (_sortedRanges) is built by sorting this by Start. Ranges
        // must still be non-overlapping; that invariant is not enforced.
        // Coverage focuses on what realistically appears in Office documents; obscure scripts
        // and historic blocks are omitted.
        private static readonly Range[] _sourceRanges = new Range[]
        {
            // Latin
            new Range(0x0020, 0x007F, UnicodeScript.Latin),        // Basic Latin (ASCII)
            new Range(0x00A0, 0x024F, UnicodeScript.Latin),        // Latin-1 Supplement + Latin Extended-A/B
            new Range(0x1E00, 0x1EFF, UnicodeScript.Latin),        // Latin Extended Additional
            new Range(0x2C60, 0x2C7F, UnicodeScript.Latin),        // Latin Extended-C

            // Greek
            new Range(0x0370, 0x03FF, UnicodeScript.Greek),        // Greek and Coptic
            new Range(0x1F00, 0x1FFF, UnicodeScript.Greek),        // Greek Extended

            // Cyrillic
            new Range(0x0400, 0x04FF, UnicodeScript.Cyrillic),     // Cyrillic
            new Range(0x0500, 0x052F, UnicodeScript.Cyrillic),     // Cyrillic Supplement
            new Range(0x2DE0, 0x2DFF, UnicodeScript.Cyrillic),     // Cyrillic Extended-A
            new Range(0xA640, 0xA69F, UnicodeScript.Cyrillic),     // Cyrillic Extended-B

            // Hebrew
            new Range(0x0590, 0x05FF, UnicodeScript.Hebrew),       // Hebrew

            // Arabic
            new Range(0x0600, 0x06FF, UnicodeScript.Arabic),       // Arabic
            new Range(0x0750, 0x077F, UnicodeScript.Arabic),       // Arabic Supplement
            new Range(0x08A0, 0x08FF, UnicodeScript.Arabic),       // Arabic Extended-A
            new Range(0xFB50, 0xFDFF, UnicodeScript.Arabic),       // Arabic Presentation Forms-A
            new Range(0xFE70, 0xFEFF, UnicodeScript.Arabic),       // Arabic Presentation Forms-B

            // Devanagari
            new Range(0x0900, 0x097F, UnicodeScript.Devanagari),   // Devanagari

            // Thai
            new Range(0x0E00, 0x0E7F, UnicodeScript.Thai),         // Thai

            // Currency
            new Range(0x20A0, 0x20CF, UnicodeScript.Currency),     // Currency Symbols

            // Math
            new Range(0x2200, 0x22FF, UnicodeScript.Math),         // Mathematical Operators
            new Range(0x27C0, 0x27EF, UnicodeScript.Math),         // Miscellaneous Mathematical Symbols-A
            new Range(0x2980, 0x29FF, UnicodeScript.Math),         // Miscellaneous Mathematical Symbols-B
            new Range(0x2A00, 0x2AFF, UnicodeScript.Math),         // Supplemental Mathematical Operators
            new Range(0x1D400, 0x1D7FF, UnicodeScript.Math),       // Mathematical Alphanumeric Symbols

            // Symbol (box drawing, geometric shapes, dingbats, misc symbols)
            new Range(0x2500, 0x257F, UnicodeScript.Symbol),       // Box Drawing
            new Range(0x2580, 0x259F, UnicodeScript.Symbol),       // Block Elements
            new Range(0x25A0, 0x25FF, UnicodeScript.Symbol),       // Geometric Shapes
            new Range(0x2700, 0x27BF, UnicodeScript.Symbol),       // Dingbats

            // CJK Han (Unified Ideographs)
            //
            // Two extra ranges below are pragmatically classified as Han even though they are
            // shared across Chinese, Japanese, and Korean writing:
            //   * CJK Symbols and Punctuation (U+3000-U+303F) — covers the ideographic full
            //     stop, ideographic comma, fullwidth space, brackets, etc.
            //   * Halfwidth and Fullwidth Forms (U+FF00-U+FFEF) — covers the fullwidth comma,
            //     fullwidth ASCII variants, halfwidth Katakana, halfwidth Hangul Jamo, etc.
            // These ranges are shared CJK punctuation/forms, not specifically Han characters.
            // In a single-language document a more accurate classification would route them
            // to the document's primary script (Hiragana/Katakana for Japanese, Hangul for
            // Korean, Han for Chinese). But without language detection we cannot do that.
            // Routing them to Han is acceptable in practice because all major CJK fonts —
            // Yu Gothic, Microsoft YaHei, Malgun Gothic, etc. — contain the same glyphs for
            // these code points.
            new Range(0x3000, 0x303F, UnicodeScript.Han),          // CJK Symbols and Punctuation
            new Range(0x3400, 0x4DBF, UnicodeScript.Han),          // CJK Unified Ideographs Extension A
            new Range(0x4E00, 0x9FFF, UnicodeScript.Han),          // CJK Unified Ideographs
            new Range(0xF900, 0xFAFF, UnicodeScript.Han),          // CJK Compatibility Ideographs
            new Range(0xFF00, 0xFFEF, UnicodeScript.Han),          // Halfwidth and Fullwidth Forms
            new Range(0x20000, 0x2A6DF, UnicodeScript.Han),        // CJK Unified Ideographs Extension B
            new Range(0x2A700, 0x2B73F, UnicodeScript.Han),        // CJK Unified Ideographs Extension C
            new Range(0x2B740, 0x2B81F, UnicodeScript.Han),        // CJK Unified Ideographs Extension D
            new Range(0x2B820, 0x2CEAF, UnicodeScript.Han),        // CJK Unified Ideographs Extension E

            // Japanese Hiragana
            new Range(0x3040, 0x309F, UnicodeScript.Hiragana),     // Hiragana

            // Japanese Katakana
            new Range(0x30A0, 0x30FF, UnicodeScript.Katakana),     // Katakana
            new Range(0x31F0, 0x31FF, UnicodeScript.Katakana),     // Katakana Phonetic Extensions

            // Korean Hangul
            new Range(0x1100, 0x11FF, UnicodeScript.Hangul),       // Hangul Jamo
            new Range(0x3130, 0x318F, UnicodeScript.Hangul),       // Hangul Compatibility Jamo
            new Range(0xAC00, 0xD7AF, UnicodeScript.Hangul),       // Hangul Syllables
            new Range(0xD7B0, 0xD7FF, UnicodeScript.Hangul),       // Hangul Jamo Extended-B

            // Emoji
            new Range(0x1F300, 0x1F5FF, UnicodeScript.Emoji),      // Miscellaneous Symbols and Pictographs
            new Range(0x1F600, 0x1F64F, UnicodeScript.Emoji),      // Emoticons
            new Range(0x1F680, 0x1F6FF, UnicodeScript.Emoji),      // Transport and Map Symbols
            new Range(0x1F700, 0x1F77F, UnicodeScript.Emoji),      // Alchemical Symbols
            new Range(0x1F780, 0x1F7FF, UnicodeScript.Emoji),      // Geometric Shapes Extended
            new Range(0x1F800, 0x1F8FF, UnicodeScript.Emoji),      // Supplemental Arrows-C
            new Range(0x1F900, 0x1F9FF, UnicodeScript.Emoji),      // Supplemental Symbols and Pictographs
            new Range(0x1FA00, 0x1FA6F, UnicodeScript.Emoji),      // Chess Symbols
            new Range(0x1FA70, 0x1FAFF, UnicodeScript.Emoji)       // Symbols and Pictographs Extended-A
        };

        // Search table: copy of _sourceRanges sorted by Start. Built once at type init.
        private static readonly Range[] _sortedRanges = BuildSortedRanges();

        private static Range[] BuildSortedRanges()
        {
            var copy = new Range[_sourceRanges.Length];
            Array.Copy(_sourceRanges, copy, _sourceRanges.Length);
            Array.Sort(copy, (a, b) => a.Start.CompareTo(b.Start));
            return copy;
        }
    }
}