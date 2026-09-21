/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Public feature selection for GSUB shaping
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// GSUB (glyph substitution) features that can be requested when shaping text.
    /// Combine flags to request more than one feature, e.g.
    /// <c>GsubFeature.Liga | GsubFeature.Clig</c>.
    /// </summary>
    /// <remarks>
    /// Each flag corresponds to a 4-character OpenType feature tag that a font's GSUB table may
    /// define. A flag being set only means the feature is requested - whether it has any effect
    /// still depends on the specific font actually defining that feature, and for which glyphs.
    /// Fonts commonly use "liga" for a small set of default ligatures (fi, fl, ff, ...) and
    /// reserve "dlig"/"clig" for optional or context-dependent ones (e.g. historical ligatures
    /// like "Th" or "ct"), so requesting only <see cref="Liga"/> will not render those.
    /// </remarks>
    [Flags]
    public enum GsubFeature
    {
        /// <summary>
        /// No GSUB substitution is requested. Text is shaped as a plain, unsubstituted sequence
        /// of glyphs mapped directly from characters.
        /// </summary>
        None = 0,

        /// <summary>
        /// Standard ligatures ("liga"). Common, typographically expected ligatures that a font
        /// applies by default - typically a small fixed set such as fi, fl, ff, ffi and ffl.
        /// </summary>
        Liga = 1 << 0,

        /// <summary>
        /// Contextual ligatures ("clig"). Ligatures a font applies only in specific contexts,
        /// as opposed to unconditionally whenever the glyphs appear together.
        /// </summary>
        Clig = 1 << 1,

        /// <summary>
        /// Discretionary ligatures ("dlig"). Optional, often decorative ligatures a font offers
        /// but does not apply by default - for example historical forms like "Th" or "ct". Off
        /// unless explicitly requested, since they are a stylistic choice rather than a
        /// correctness requirement.
        /// </summary>
        Dlig = 1 << 2,

        /// <summary>
        /// Contextual alternates ("calt"). Glyph substitutions a font applies based on
        /// surrounding context rather than unconditionally. Connected or cursive script fonts
        /// commonly rely on this to join adjacent letterforms - without it, such fonts can
        /// render as visibly disconnected glyphs rather than a flowing script. Unlike
        /// <see cref="Dlig"/>, this is generally a correctness expectation rather than a
        /// stylistic choice, which is why it is included by default.
        /// </summary>
        Calt = 1 << 3,

        /// <summary>
        /// Glyph composition and decomposition ("ccmp"). Rearranges glyphs so that other
        /// features and mark positioning can work on them - most often by DECOMPOSING a
        /// precomposed glyph into a base glyph plus separate combining marks, but composition
        /// in the other direction is equally valid. The OpenType specification treats this as
        /// an always-on feature that shaping engines apply before other substitutions, rather
        /// than an optional stylistic choice, which is why it is included by default and runs
        /// first.
        /// </summary>
        Ccmp = 1 << 4
    }
}