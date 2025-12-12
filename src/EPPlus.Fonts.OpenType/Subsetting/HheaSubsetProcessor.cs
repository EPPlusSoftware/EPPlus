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
    internal class HheaSubsetProcessor : IFontSubsetProcessor
    {
        public void Process(FontSubsettingContext context)
        {
            var original = context.OriginalFont.HheaTable;
            if (original == null) return;

            var hhea = original.Clone();
            hhea.numberOfHMetrics = (ushort)context.NewToOldGlyphId.Count;
            context.SubsetFont.AddOrReplaceTable(hhea);
        }
    }
}
