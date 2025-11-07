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
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public abstract class CmapSubtableBase : FontTableElement
    {
        /// <summary>
        /// Format identifier (0, 4, 6, etc.)
        /// </summary>
        public abstract ushort Format { get; }

        /// <summary>
        /// Length of the subtable in bytes
        /// </summary>
        public abstract ushort Length { get; internal set; }

        /// <summary>
        /// Language code (optional usage depending on format)
        /// </summary>
        public abstract ushort Language { get; internal set; }
    }
}
