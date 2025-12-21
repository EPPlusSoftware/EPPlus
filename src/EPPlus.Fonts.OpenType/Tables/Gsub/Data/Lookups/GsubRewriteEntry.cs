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
namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// En hjälpstruktur för att hålla mappningen mellan gamla och nya Glyph IDs 
    /// under omskrivningen (Rewrite) av GSUB-tabeller.
    /// </summary>
    internal struct GsubRewriteEntry
    {
        /// <summary>
        /// Det nya Glyph ID:t för källtecknet (Input).
        /// </summary>
        public ushort NewInput;

        /// <summary>
        /// Det nya Glyph ID:t för ersättningstecknet (Output/Substitute).
        /// </summary>
        public ushort NewOutput;
    }
}