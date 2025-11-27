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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Scanner
{
    internal interface IScannedFont
    {
        string FontFamilyName { get; }

        string FontSubFamilyName { get; set; }

        string FilePath { get; set; }

        FontFormat Format { get; set; }

        IEnumerable<ScannedFont>? SubFonts { get; }

        long? TtcOffset { get; }

        byte[] GetTableBytes(string tag);
    }
}
