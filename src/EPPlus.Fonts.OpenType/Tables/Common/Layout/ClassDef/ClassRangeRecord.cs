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
using System;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef
{
    /// <summary>
    /// Represents a single ClassRangeRecord in a ClassDef Format 2 table.
    /// Maps a continuous glyph ID range to a class value.
    /// </summary>
    public class ClassRangeRecord
    {
        /// <summary>
        /// First glyph ID in the range (inclusive).
        /// </summary>
        public ushort StartGlyphID { get; set; }

        /// <summary>
        /// Last glyph ID in the range (inclusive).
        /// </summary>
        public ushort EndGlyphID { get; set; }

        /// <summary>
        /// Class value assigned to all glyphs in [StartGlyphID, EndGlyphID].
        /// </summary>
        public ushort Class { get; set; }
    }
}
