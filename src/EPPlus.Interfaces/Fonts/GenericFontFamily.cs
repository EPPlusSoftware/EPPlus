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
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// The CSS generic font family a font belongs to. Used as the last entry of a CSS
    /// font-family stack, so a browser that has none of the named fonts picks a font of
    /// the same kind.
    /// </summary>
    public enum GenericFontFamily
    {
        /// <summary>
        /// Proportional font without serifs. CSS keyword <c>sans-serif</c>. Also used when
        /// the font does not say what it is.
        /// </summary>
        SansSerif,
        /// <summary>
        /// Proportional font with serifs. CSS keyword <c>serif</c>.
        /// </summary>
        Serif,
        /// <summary>
        /// Font where all glyphs have the same advance width. CSS keyword <c>monospace</c>.
        /// </summary>
        Monospace
    }
}
