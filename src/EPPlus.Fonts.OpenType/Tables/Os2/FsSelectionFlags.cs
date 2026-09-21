/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2026         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System;

namespace EPPlus.Fonts.OpenType.Tables.Os2
{
    [Flags]
    public enum FsSelectionFlags : ushort
    {
        Italic = 1 << 0,            // Bit 0
        Underscore = 1 << 1,        // Bit 1
        Negative = 1 << 2,          // Bit 2
        Outlined = 1 << 3,          // Bit 3
        Strikeout = 1 << 4,         // Bit 4
        Bold = 1 << 5,              // Bit 5
        Regular = 1 << 6,           // Bit 6
        UseTypoMetrics = 1 << 7,    // Bit 7
        WWS = 1 << 8,               // Bit 8
        Oblique = 1 << 9            // Bit 9
                                    // Bits 10-15 are reserved
    }
}
