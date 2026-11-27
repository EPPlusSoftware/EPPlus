/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/23/2026         EPPlus Software AB           Unhinted subsets for web output
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Options for creating a font subset.
    /// </summary>
    internal sealed class SubsetOptions
    {
        /// <summary>
        /// Prefixes the names of the subset with a six-letter subset tag, e.g. "ABCDEF+Roboto".
        /// </summary>
        internal bool AddSubsetTag { get; set; }

        /// <summary>
        /// Removes TrueType hinting so the subset is a valid unhinted font: the hinting tables
        /// (cvt, fpgm, prep), the instructions of every glyph, and the hinting limits in maxp.
        /// Meant for web output, where browsers on most platforms ignore hinting and text is laid out
        /// with geometric precision and a fixed textLength. False keeps hinting, as PDF export does.
        /// The tables and the glyph instructions are always removed together: instructions that call
        /// functions in a missing fpgm make the font invalid.
        /// </summary>
        internal bool RemoveHinting { get; set; }
    }
}