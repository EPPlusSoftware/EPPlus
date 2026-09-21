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
    public enum FsTypeFlags : ushort
    {
        /// <summary>Installable embedding (no restrictions). Bits 0-3 all clear.</summary>
        Installable = 0x0000,
        /// <summary>Restricted License embedding. Bit 1.</summary>
        RestrictedLicense = 0x0002,
        /// <summary>Preview &amp; Print embedding. Bit 2.</summary>
        PreviewPrint = 0x0004,
        /// <summary>Editable embedding. Bit 3.</summary>
        Editable = 0x0008,
        /// <summary>No subsetting: font must be embedded whole, not subsetted. Bit 8.</summary>
        NoSubsetting = 0x0100,
        /// <summary>Bitmap embedding only: only bitmap data may be embedded. Bit 9.</summary>
        BitmapEmbeddingOnly = 0x0200,
    }
}
