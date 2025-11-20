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
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable4 : CmapSubtableBase
    {

        public override ushort Format { get; } = 4;

        public override uint Length { get; internal set; }

        public override  uint Language { get; internal set; }

        public ushort SegCountX2 { get; internal set; }
        public ushort SearchRange { get; internal set; }
        public ushort EntrySelector { get; internal set; }
        public ushort RangeShift { get; internal set; }

        public ushort[] EndCode { get; internal set; } = new ushort[0];
        public ushort ReservedPad { get; internal set; }
        public ushort[] StartCode { get; internal set; } = new ushort[0];
        public short[] IdDelta { get; internal set; } = new short[0];
        public ushort[] IdRangeOffset { get; internal set; } = new ushort[0];

        public ushort[] GlyphIdArray { get; internal set; } = new ushort[0];

        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            int segCount = EndCode.Length;

            for (int i = 0; i < segCount; i++)
            {
                ushort startCode = StartCode[i];
                ushort endCode = EndCode[i];
                short idDelta = IdDelta[i];
                ushort idRangeOffset = IdRangeOffset[i];

                for (uint charCode = startCode; charCode <= endCode; charCode++)
                {
                    ushort glyphIndex;

                    if (idRangeOffset == 0)
                    {
                        glyphIndex = (ushort)((charCode + idDelta) % 65536);
                    }
                    else
                    {
                        int offsetIndex = (idRangeOffset / 2) + (int)(charCode - startCode) - (segCount - i);
                        if (offsetIndex >= 0 && offsetIndex < GlyphIdArray.Length)
                        {
                            ushort glyphId = GlyphIdArray[offsetIndex];
                            if (glyphId != 0)
                            {
                                glyphIndex = (ushort)((glyphId + idDelta) % 65536);
                            }
                            else
                            {
                                glyphIndex = 0;
                            }
                        }
                        else
                        {
                            glyphIndex = 0;
                        }
                    }

                    if (glyphIndex != 0)
                    {
                        mapping.AddMapping(charCode, glyphIndex);
                    }
                }
            }

            return mapping;
        }

        internal override int MapCodePointToGlyph(int codePoint)
        {
            var segCount = EndCode.Length;
            for (int i = 0; i < segCount; i++)
            {
                if (codePoint >= StartCode[i] && codePoint <= EndCode[i])
                {
                    if (IdRangeOffset[i] == 0)
                    {
                        return (codePoint + IdDelta[i]) & 0xFFFF;
                    }
                    else
                    {
                        int offset = IdRangeOffset[i] / 2 + (codePoint - StartCode[i]) - (segCount - i);
                        if (offset >= 0 && offset < GlyphIdArray.Length)
                        {
                            return GlyphIdArray[offset];
                        }
                    }
                }
            }
            return -1; // Not found

        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable4Serializer();
            serializer.Serialize(this, writer);
        }

        public void BuildFromMappings(List<CharGlyphMapping> mappings)
        {
            if (mappings.Count == 0) return;

            mappings = mappings.OrderBy(m => m.CharCode).ToList();

            var startCodes = new List<ushort>();
            var endCodes = new List<ushort>();
            var deltas = new List<short>();
            var rangeOffsets = new List<ushort>();

            foreach (var map in mappings)
            {
                startCodes.Add((ushort)map.CharCode);
                endCodes.Add((ushort)map.CharCode);
                deltas.Add((short)(map.GlyphId - map.CharCode));
                rangeOffsets.Add(0);
            }

            // Terminator segment
            startCodes.Add(0xFFFF);
            endCodes.Add(0xFFFF);
            deltas.Add(0); // safer than 1
            rangeOffsets.Add(0);

            StartCode = startCodes.ToArray();
            EndCode = endCodes.ToArray();
            IdDelta = deltas.ToArray();
            IdRangeOffset = rangeOffsets.ToArray();
            ReservedPad = 0;

            int segCount = startCodes.Count;
            SegCountX2 = (ushort)(segCount * 2);

            int log2 = (int)Math.Floor(Math.Log(segCount, 2));
            ushort maxPower = (ushort)(1 << log2);
            SearchRange = (ushort)(maxPower * 2);
            EntrySelector = (ushort)log2;
            RangeShift = (ushort)(SegCountX2 - SearchRange);

            GlyphIdArray = new ushort[0];

            Length = (ushort)(16 + segCount * 8 + 2 + GlyphIdArray.Length * 2);
        }
    }
}
