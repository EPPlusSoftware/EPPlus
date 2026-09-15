/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Public feature selection for GPOS shaping
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// GPOS (glyph positioning) features that can be requested when shaping text.
    /// Combine flags to request more than one feature, e.g.
    /// <c>GposFeature.Kern | GposFeature.Mark</c>.
    /// </summary>
    /// <remarks>
    /// Each flag corresponds to a 4-character OpenType feature tag that a font's GPOS table may
    /// define. A flag being set only means the feature is requested - whether it has any effect
    /// still depends on the specific font actually defining that feature, and for which glyphs
    /// and scripts.
    /// </remarks>
    [Flags]
    public enum GposFeature
    {
        /// <summary>
        /// No GPOS positioning is requested. Glyphs are placed using only their default
        /// advance widths, with no pair kerning or mark attachment applied.
        /// </summary>
        None = 0,

        /// <summary>
        /// Pair kerning ("kern"). Adjusts the spacing between specific pairs of glyphs (e.g. "AV",
        /// "To") so they sit visually closer or further apart than their default advance widths
        /// alone would produce.
        /// </summary>
        Kern = 1 << 0,

        /// <summary>
        /// Mark-to-base attachment ("mark"). Positions combining marks (accents, diacritics)
        /// relative to the base glyph they attach to, rather than at the base glyph's default
        /// advance position. Needed for decomposed sequences such as a base letter followed by a
        /// combining accent (e.g. "A" + U+0301) to render correctly.
        /// </summary>
        Mark = 1 << 1
    }
}