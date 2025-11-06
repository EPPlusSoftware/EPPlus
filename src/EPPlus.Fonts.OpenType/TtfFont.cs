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
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Kern;
using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Represents an open type font
    /// </summary>
    [DebuggerDisplay("{FullName} {SubFamily}")]
    public class TtfFont : OpenTypeFont
    {

        internal TtfFont(FontsBinaryReader reader, FontFormat format)
            : this(reader, -1, format)
        {
        }

        internal TtfFont(FontsBinaryReader reader, long startOffset, FontFormat format) : base(reader, startOffset, format)
        {
            _glyphTableLoader = TableLoaders.GetGlyphTableLoader(tblSettings);
            _kernTableLoader = TableLoaders.GetKernTableLoader(tblSettings);
        }
    }
}
