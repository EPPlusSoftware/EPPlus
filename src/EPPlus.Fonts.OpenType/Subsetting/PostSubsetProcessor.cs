using EPPlus.Fonts.OpenType.Tables.Post;
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

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Creates a minimal, correct 'post' table for the subset font.
    /// Uses format 3.0 (no glyph names) – the recommended format for subset fonts.
    /// Preserves italicAngle, underlinePosition, underlineThickness and isFixedPitch from the original.
    /// .NET 3.5 compatible.
    /// </summary>
    internal class PostSubsetProcessor : IFontSubsetProcessor
    {
        public void Process(FontSubsettingContext context)
        {
            var originalFont = context.OriginalFont;
            var subsetFont = context.SubsetFont;

            if (originalFont.PostTable == null)
                return; // No post table in source font

            var originalPost = originalFont.PostTable;

            // Create a clean format 3.0 post table (no glyph names – allowed and preferred for subsets)
            var post = new PostTable
            {
                // Format 3.0 in fixed 16.16 representation
                version = new Version16Dot16(0x00030000),

                // Preserve typographic behaviour from the original font
                italicAngle = originalPost.italicAngle,
                underlinePosition = originalPost.underlinePosition,
                underlineThickness = originalPost.underlineThickness,
                isFixedPitch = originalPost.isFixedPitch,

                // These fields are only used in format 2.0 – set to zero
                minMemType42 = 0,
                maxMemType42 = 0,
                minMemType1 = 0,
                maxMemType1 = 0,

                // Explicitly clear format 2.0 data (defensive programming)
                glyphNameIndex = null,
                glyphNames = null
            };

            subsetFont.AddOrReplaceTable(post);
        }
    }
}
