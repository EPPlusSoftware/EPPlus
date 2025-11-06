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
using EPPlus.Fonts.OpenType.Tables.Head;
using System;

namespace EPPlus.Fonts.OpenType.Tables.Loca
{
    /// <summary>
    /// The indexToLoc table stores the offsets to the locations of the glyphs in the font, relative to the beginning of the glyphData table
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/loca
    /// </summary>
    public class LocaTable : FontTableBase
    {
        public uint[] Offsets { get; set; }

        public HeadTable.IndexToLocFormats IndexToLocFormat { get; set; }


        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            if (IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                foreach (var offset in Offsets)
                {
                    ushort shortOffset = (ushort)(offset / 2);
                    writer.WriteUInt16BigEndian(shortOffset);
                }
            }
            else if (IndexToLocFormat == HeadTable.IndexToLocFormats.Offset32)
            {
                foreach (var offset in Offsets)
                {
                    writer.WriteUInt32BigEndian(offset);
                }
            }
            else
            {
                throw new InvalidOperationException("Unsupported IndexToLocFormat.");
            }
        }

    }
}
