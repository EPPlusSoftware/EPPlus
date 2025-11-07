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
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable4Serializer : CmapSubtableSerializerBase<CmapSubtable4>
    {
        internal override void Serialize(CmapSubtable4 subTable, FontsBinaryWriter writer)
        {
            // Clone the segment list and add the required sentinel segment
            var segments = new List<CmapSubtable4Segment>(subTable.Segments)
        {
            new CmapSubtable4Segment
            {
                StartCode = 0xFFFF,
                EndCode = 0xFFFF,
                IdDelta = 1,
                IdRangeOffset = 0,
                GlyphIdArray = null
            }
        };

            int segCount = segments.Count;
            int segCountX2 = segCount * 2;

            // Calculate search parameters
            int power = (int)Math.Floor(Math.Log(segCount, 2));
            ushort searchRange = (ushort)(Math.Pow(2, power) * 2);
            ushort entrySelector = (ushort)power;
            ushort rangeShift = (ushort)(segCountX2 - searchRange);

            // Debug output: segment info
            Debug.WriteLine($"Segment count (incl. sentinel): {segCount}");
            Debug.WriteLine($"SearchRange: {searchRange}, EntrySelector: {entrySelector}, RangeShift: {rangeShift}");

            // Prepare arrays
            ushort[] endCodes = segments.Select(s => s.EndCode).ToArray();
            ushort[] startCodes = segments.Select(s => s.StartCode).ToArray();
            short[] idDeltas = segments.Select(s => s.IdDelta).ToArray();
            ushort[] idRangeOffsets = new ushort[segCount];
            List<ushort> glyphIdArray = new();

            // Calculate idRangeOffsets and build glyphIdArray
            int offsetBase = 16 + segCount * 8 + 2; // Header + 4 arrays + reservedPad
            for (int i = 0; i < segCount; i++)
            {
                var segment = segments[i];
                if (segment.GlyphIdArray != null && segment.GlyphIdArray.Length > 0)
                {
                    int offset = (segCount - i) * 2 + glyphIdArray.Count * 2;
                    idRangeOffsets[i] = (ushort)offset;
                    glyphIdArray.AddRange(segment.GlyphIdArray);
                }
                else
                {
                    idRangeOffsets[i] = 0;
                }

                Debug.WriteLine($"Segment {i}: Start={segment.StartCode}, End={segment.EndCode}, Delta={segment.IdDelta}, Offset={idRangeOffsets[i]}, Glyphs={(segment.GlyphIdArray?.Length ?? 0)}");
            }

            Debug.WriteLine($"Total glyphIdArray entries: {glyphIdArray.Count}");

            // Calculate total length
            int length = offsetBase + glyphIdArray.Count * 2;
            Debug.WriteLine($"Calculated Format 4 subtable length: {length}");

            // Write header
            writer.WriteUInt16BigEndian(subTable.Format);        // format = 4
            writer.WriteUInt16BigEndian((ushort)length);
            writer.WriteUInt16BigEndian(subTable.Language);
            writer.WriteUInt16BigEndian((ushort)segCountX2);
            writer.WriteUInt16BigEndian(searchRange);
            writer.WriteUInt16BigEndian(entrySelector);
            writer.WriteUInt16BigEndian(rangeShift);

            // Write segment arrays
            foreach (var endCode in endCodes)
                writer.WriteUInt16BigEndian(endCode);

            writer.WriteUInt16BigEndian(0); // reservedPad

            foreach (var startCode in startCodes)
                writer.WriteUInt16BigEndian(startCode);

            foreach (var delta in idDeltas)
                writer.WriteInt16BigEndian(delta);

            foreach (var offset in idRangeOffsets)
                writer.WriteUInt16BigEndian(offset);

            // Write glyphIdArray
            foreach (var glyphId in glyphIdArray)
                writer.WriteUInt16BigEndian(glyphId);
        }
    }
}
