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
    internal class MaxpSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            // No implementation
        }

        public void Rewrite(FontSubsettingContext context)
        {
            var original = context.OriginalFont.MaxpTable;
            if (original == null) return;

            var maxp = original.Clone();
            maxp.numGlyphs = (ushort)context.NewToOldGlyphId.Count;

            if (context.RemoveHinting)
            {
                // No instructions remain, so nothing is reserved for them. maxZones is 1 rather than 0:
                // 1 means no twilight zone, and 0 is not a valid value.
                maxp.maxZones = 1;
                maxp.maxTwilightPoints = 0;
                maxp.maxStorage = 0;
                maxp.maxFunctionDefs = 0;
                maxp.maxInstructionDefs = 0;
                maxp.maxStackElements = 0;
                maxp.maxSizeOfInstructions = 0;
            }

            context.SubsetFont.AddOrReplaceTable(maxp);
        }
    }
}
