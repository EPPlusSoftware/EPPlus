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
    public interface IFontSubsetProcessor
    {
        /// <summary>
        /// Phase 1: Analyzes the original font to discover dependencies (e.g., ligature glyphs).
        /// </summary>
        void Discover(FontSubsettingContext context);


        /// <summary>
        /// Phase 2: Creates the new subsetted table based on the discovered dependencies.
        /// </summary>
        public void Rewrite(FontSubsettingContext context);
    }
}
