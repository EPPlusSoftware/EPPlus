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
    internal class HeadSubsetProcessor : IFontSubsetProcessor
    {
        public void Discover(FontSubsettingContext context)
        {
            var original = context.OriginalFont.HeadTable;
            if (original == null) return;

            var head = original.Clone();
            // checkSumAdjustment will be recalculated at font save – we leave it for now
            context.SubsetFont.AddOrReplaceTable(head);
        }

        public void Rewrite(FontSubsettingContext context)
        {
            // No implementation
        }
    }
}
