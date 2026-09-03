/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/20/2026         EPPlus Software AB           Single-font subsetter extracted from FontSubsetManager
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Utils;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Subsets one font down to a given set of code points. This is a low-level building block:
    /// it does not resolve fallback chains and makes no embedding-policy decisions — the caller
    /// owns all of that. Kept separate so it can be unit-tested in isolation and reused by any
    /// component that needs to reduce a single font.
    /// </summary>
    internal sealed class SingleFontSubsetter
    {
        /// <summary>
        /// Produces a subset of <paramref name="font"/> containing only the glyphs required for
        /// <paramref name="codePoints"/>. Returns the font unchanged when it is already a subset
        /// or when no code points are supplied. If subsetting fails, the original font is returned
        /// so the caller always receives an embeddable instance.
        /// </summary>
        public OpenTypeFont Subset(OpenTypeFont font, HashSet<int> codePoints)
        {
            if (font == null)
                throw new ArgumentNullException("font");

            if (font.IsSubset || codePoints == null || codePoints.Count == 0)
                return font;

            try
            {
                var chars = CodePointUtil.CodePointsToString(codePoints);
                return font.CreateSubset(chars);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "Warning: could not subset '" +
                    (font.NameTable != null ? font.NameTable.GetFullFontName() : "(unknown)") +
                    "': " + ex.Message);
                return font;
            }
        }
    }
}