/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/17/2026         EPPlus Software AB           Pass-through of hinting tables
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;

namespace EPPlus.Fonts.OpenType.Subsetting.Processors
{
    /// <summary>
    /// Copies tables that do not reference glyph IDs unchanged from the original font to the subset.
    /// Glyph instructions in 'glyf' call functions in 'fpgm' and read values from 'cvt ', so these
    /// tables must be kept for the hinted glyphs to execute correctly in rasterizers and pass OTS.
    /// Tables missing in the original are not added.
    /// .NET 3.5 compatible.
    /// </summary>
    internal class PassThroughTableSubsetProcessor : IFontSubsetProcessor
    {
        private static readonly string[] PassThroughTags = new string[]
        {
            TableNames.Cvt,
            TableNames.Fpgm,
            TableNames.Prep,
            TableNames.Gasp
        };

        public void Discover(FontSubsettingContext context)
        {
            // No glyph dependencies
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var original = context.OriginalFont;
            foreach (var tag in PassThroughTags)
            {
                if (!original.TableRecords.ContainsKey(tag))
                    continue;

                var data = original.GetTableData(tag);
                if (data == null)
                    continue;

                context.SubsetFont.AddOrReplaceTable(new RawFontTable(tag, data));
            }
        }
    }
}