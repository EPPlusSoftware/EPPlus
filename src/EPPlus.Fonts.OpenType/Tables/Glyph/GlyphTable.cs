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
namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    /// <summary>
    /// This table contains information that describes the glyphs in the font in the TrueType outline format
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/glyf
    /// </summary>
    public class GlyphTable
    {
        public GlyphHeader[] Glyphs { get; set; }
    }
}
