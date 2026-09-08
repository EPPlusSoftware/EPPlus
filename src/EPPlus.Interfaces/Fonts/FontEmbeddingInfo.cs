/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2026         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Information passed to the <see cref="IEpplusFontConfiguration.OnFontEmbedding"/>
    /// callback so the caller can decide how a font should be embedded.
    /// </summary>
    public class FontEmbeddingInfo
    {
        public FontEmbeddingInfo(string fontName, FontEmbeddingRestriction restriction)
        {
            FontName = fontName;
            Restriction = restriction;
        }

        /// <summary>The full name of the font being prepared for embedding.</summary>
        public string FontName { get; private set; }

        /// <summary>The restriction the font declares via its OS/2 fsType field.</summary>
        public FontEmbeddingRestriction Restriction { get; private set; }
    }
}
