/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           vhea subset processor (vertical text support)
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Subsetting.Processors
{
    /// <summary>
    /// Creates the subsetted 'vhea' (Vertical Header) table.
    /// Only runs if the original font contains a vhea table.
    /// Analogous to <see cref="HheaSubsetProcessor"/> for the horizontal header.
    /// Clones the original vhea and updates NumberOfVMetrics to match
    /// the subset glyph count (mirroring how HheaSubsetProcessor updates numberOfHMetrics).
    /// </summary>
    internal class VheaSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            // No additional glyphs to discover
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var originalVhea = context.OriginalFont.VheaTable;

            // vhea is optional - skip silently if not present
            if (originalVhea == null) return;

            var vhea = originalVhea.Clone();
            vhea.NumberOfVMetrics = (ushort)context.NewToOldGlyphId.Count;
            context.SubsetFont.AddOrReplaceTable(vhea);
        }
    }
}