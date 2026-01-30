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

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage
{
    public class CoverageRangeRecord : FontTableElement
    {
        public ushort StartGlyphID { get; set; }
        public ushort EndGlyphID { get; set; }
        public ushort StartCoverageIndex { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
