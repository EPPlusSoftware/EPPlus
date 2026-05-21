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
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Unicode scripts used for per-script glyph-level fallback. Each value corresponds
    /// to a Unicode block or group of blocks that share a writing system, and is the key
    /// for configuring fallback fonts via
    /// <c>EpplusFontConfiguration.SetScriptFallback(UnicodeScript, params string[])</c>.
    ///
    /// Only scripts that are realistic in Office documents are included. Adding more is
    /// straightforward but requires updating <c>UnicodeScriptClassifier</c> as well.
    /// </summary>
    public enum UnicodeScript
    {
        /// <summary>The character does not belong to any of the supported scripts.</summary>
        Unknown = 0,

        /// <summary>Latin script (ASCII, Latin-1, Latin Extended). Covers English, most European languages.</summary>
        Latin,

        /// <summary>Cyrillic script. Covers Russian, Ukrainian, Bulgarian, Serbian, and other Slavic languages.</summary>
        Cyrillic,

        /// <summary>Greek script. Covers Modern and Polytonic Greek.</summary>
        Greek,

        /// <summary>Arabic script. Covers Arabic, Persian, Urdu, and related languages.</summary>
        Arabic,

        /// <summary>Hebrew script.</summary>
        Hebrew,

        /// <summary>Thai script.</summary>
        Thai,

        /// <summary>Devanagari script. Covers Hindi, Marathi, Nepali, and many other Indian languages.</summary>
        Devanagari,

        /// <summary>CJK Unified Ideographs (Han characters). Used in Chinese, Japanese, Korean, Vietnamese.</summary>
        Han,

        /// <summary>Japanese Hiragana syllabary.</summary>
        Hiragana,

        /// <summary>Japanese Katakana syllabary.</summary>
        Katakana,

        /// <summary>Korean Hangul script (syllables and Jamo).</summary>
        Hangul,

        /// <summary>Emoji and pictographic symbols (U+1F300-U+1FAFF and related blocks).</summary>
        Emoji,

        /// <summary>Mathematical symbols and operators (U+2200-U+22FF, U+27C0-U+27EF, U+2A00-U+2AFF).</summary>
        Math,

        /// <summary>Currency symbols (U+20A0-U+20CF).</summary>
        Currency,

        /// <summary>Box drawing, block elements, and miscellaneous symbols used in tables and diagrams.</summary>
        Symbol
    }
}