/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/19/2026         EPPlus Software AB           Generic font family for CSS font stacks
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Decides the CSS generic family of a font from its own tables.
    ///
    /// In order:
    ///   1. post.isFixedPitch set              → monospace
    ///   2. OS/2 sFamilyClass, class 1-7       → serif (the IBM classes for serif designs)
    ///   3. PANOSE, Latin text (family type 2):
    ///        proportion 9 (monospaced)        → monospace
    ///        serif style 2-10                 → serif
    ///   4. anything else                      → sans-serif
    ///
    /// Many fonts, Google Fonts in particular, leave both sFamilyClass and PANOSE at 0. Such
    /// fonts are reported as sans-serif even when they have serifs, since nothing in the font
    /// says otherwise.
    /// </summary>
    internal static class GenericFontFamilyClassifier
    {
        private const int PanoseFamilyTypeLatinText = 2;
        private const int PanoseProportionMonospaced = 9;

        internal static GenericFontFamily Classify(OpenTypeFont font)
        {
            if (font == null)
                return GenericFontFamily.SansSerif;

            var post = font.PostTable;
            if (post != null && post.isFixedPitch != 0)
                return GenericFontFamily.Monospace;

            var os2 = font.Os2Table;
            if (os2 == null)
                return GenericFontFamily.SansSerif;

            // The class id is the high byte; the low byte is the subclass.
            var familyClass = (os2.sFamilyClass >> 8) & 0xFF;
            if (familyClass >= 1 && familyClass <= 7)
                return GenericFontFamily.Serif;

            var panose = os2.panose;
            if (panose != null && panose.Length >= 4 && panose[0] == PanoseFamilyTypeLatinText)
            {
                if (panose[3] == PanoseProportionMonospaced)
                    return GenericFontFamily.Monospace;

                var serifStyle = panose[1];
                if (serifStyle >= 2 && serifStyle <= 10)
                    return GenericFontFamily.Serif;
            }

            return GenericFontFamily.SansSerif;
        }
    }
}
