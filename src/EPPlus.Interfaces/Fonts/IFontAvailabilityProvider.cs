/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Optional capability for font resolvers that can report on font availability
    /// without performing a full font load. Implement this on a custom
    /// <see cref="IFontResolver"/> to support
    /// <c>OpenTypeFonts.GetFontAvailability(...)</c> with full fidelity — including
    /// distinguishing between an exact subfamily match and a family-only match.
    ///
    /// Resolvers that do not implement this interface still work; in that case
    /// <c>OpenTypeFonts.GetFontAvailability</c> falls back to probing via
    /// <see cref="IFontResolver.ResolveFont"/>, which can only distinguish
    /// "found" from "not found".
    /// </summary>
    public interface IFontAvailabilityProvider
    {
        /// <summary>
        /// Checks whether the specified font is available, and at what level
        /// of specificity.
        /// </summary>
        /// <param name="fontName">The font family name to check.</param>
        /// <param name="subFamily">The requested subfamily (Regular, Bold, Italic, BoldItalic).</param>
        FontAvailability GetFontAvailability(string fontName, FontSubFamily subFamily);
    }
}
