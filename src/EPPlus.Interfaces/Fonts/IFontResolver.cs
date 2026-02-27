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
  02/26/2026         EPPlus Software AB           Simplified to return raw font bytes
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Resolves a font by name and subfamily to raw font bytes.
    /// Implement this interface to provide fonts from any source (file system, database, embedded resources, etc.).
    /// Font format (TTF/OTF) is detected automatically from the returned bytes.
    /// </summary>
    public interface IFontResolver
    {
        /// <summary>
        /// Resolves a font to its raw bytes.
        /// </summary>
        /// <param name="fontName">Font family name (e.g. "Roboto")</param>
        /// <param name="subFamily">Font subfamily (Regular, Bold, Italic, etc.)</param>
        /// <returns>Raw font bytes, or null if the font could not be found</returns>
        byte[] ResolveFont(string fontName, FontSubFamily subFamily);
    }
}