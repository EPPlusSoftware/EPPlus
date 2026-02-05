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
  01/19/2026         EPPlus Software AB           Performance optimization with binary search
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable4 : CmapSubtableBase
    {

        public override ushort Format { get; } = 4;

        public override uint Length { get; internal set; }

        public override uint Language { get; internal set; }

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
            // Performance optimization: Use binary search to find the segment
            // EndCode array is sorted in ascending order per OpenType spec

            if (codePoint < 0 || codePoint > 0xFFFF)
                return -1;

            int segCount = EndCode.Length;
            if (segCount == 0)
                return -1;

            // Binary search for the segment containing this codePoint
            // We're looking for the first EndCode >= codePoint
            int left = 0;
            int right = segCount - 1;
            int segmentIndex = -1;

            while (left <= right)
            {
                int mid = left + (right - left) / 2;

                if (EndCode[mid] >= codePoint)
                {
                    segmentIndex = mid;
                    right = mid - 1;  // Continue searching left for earlier match
                }
                else
                {
                    left = mid + 1;
                }
            }

            // If no segment found or codePoint is before the segment's start, return -1
            if (segmentIndex == -1 || codePoint < StartCode[segmentIndex])
                return -1;

            // Found the segment, now map to glyph
            int i = segmentIndex;

            if (IdRangeOffset[i] == 0)
            {
                // Simple offset mapping
                return (codePoint + IdDelta[i]) & 0xFFFF;
            }
            else
            {
                // Index into GlyphIdArray
                int offset = IdRangeOffset[i] / 2 + (codePoint - StartCode[i]) - (segCount - i);

                if (offset >= 0 && offset < GlyphIdArray.Length)
                {
                    ushort glyphId = GlyphIdArray[offset];
                    if (glyphId != 0)
                    {
                        return (glyphId + IdDelta[i]) & 0xFFFF;
                    }
                }
            }

            return -1;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable4Serializer();
            serializer.Serialize(this, writer);
        }
    }
}