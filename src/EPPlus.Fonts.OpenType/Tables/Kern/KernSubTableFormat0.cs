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

namespace EPPlus.Fonts.OpenType.Tables.Kern
{
    public class KernSubTableFormat0 : FontTableElement
    {
        internal KernSubTableFormat0(FontsBinaryReader reader)
        {
            nPairs = reader.ReadUInt16BigEndian();
            SearchRange = reader.ReadUInt16BigEndian();
            EntrySelector = reader.ReadUInt16BigEndian();
            RangeShift = reader.ReadUInt16BigEndian();
        }

        public ushort nPairs { get; set; }

        public ushort SearchRange { get; set; }
        public ushort EntrySelector { get; set; }
        public ushort RangeShift { get; set; }

        public KerningPair[] Pairs { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {


            writer.WriteUInt16BigEndian(nPairs);

            // Beräkna sökparametrar
            ushort maxPowerOf2 = 1;
            while (maxPowerOf2 * 2 <= nPairs)
                maxPowerOf2 *= 2;

            ushort searchRange = (ushort)(maxPowerOf2 * 6);
            ushort entrySelector = (ushort)(Math.Log(maxPowerOf2) / Math.Log(2));
            ushort rangeShift = (ushort)((nPairs * 6) - searchRange);

            writer.WriteUInt16BigEndian(searchRange);
            writer.WriteUInt16BigEndian(entrySelector);
            writer.WriteUInt16BigEndian(rangeShift);

            // Skriv kerningpar
            foreach (var pair in Pairs)
            {
                writer.WriteUInt16BigEndian(pair.left);
                writer.WriteUInt16BigEndian(pair.right);
                writer.WriteInt16BigEndian(pair.value);
            }

        }
    }
}
