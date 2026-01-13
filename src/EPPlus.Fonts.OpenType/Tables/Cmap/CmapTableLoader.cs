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
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapTableLoader : TableLoader<CmapTable>
    {
        public CmapTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Cmap)
        {
        }

        protected override CmapTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;

            var table = new CmapTable
            {
                Version = _reader.ReadUInt16BigEndian(),
                NumTables = _reader.ReadUInt16BigEndian()
            };

            // Read encoding records
            for (var x = 0; x < table.NumTables; x++)
            {
                var enc = new EncodingRecord(_reader);
                table.EncodingRecords.Add(enc);
            }

            // Deduplicate subtables by offset
            var subtableCache = new Dictionary<uint, CmapSubtableBase>();

            // Process subtables
            for (var x = 0; x < table.NumTables; x++)
            {
                var enc = table.EncodingRecords[x];
                var currentPos = _offset + enc.SubtableOffset;

                if (subtableCache.TryGetValue(enc.SubtableOffset, out var existingSubtable))
                {
                    // Reuse existing subtable
                    enc.Subtable = existingSubtable;
                    continue;
                }

                // Read format to determine deserializer
                long savedPos = _reader.BaseStream.Position;
                _reader.BaseStream.Position = currentPos;
                var format = _reader.ReadUInt16BigEndian();
                _reader.BaseStream.Position = savedPos;

                // Deserialize based on format
                switch (format)
                {
                    case 0:
                        var sub0 = new CmapSubtable0Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub0);
                        subtableCache[enc.SubtableOffset] = sub0;
                        enc.Subtable = sub0;
                        break;

                    case 4:
                        var sub4 = new CmapSubtable4Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub4);
                        subtableCache[enc.SubtableOffset] = sub4;
                        enc.Subtable = sub4;
                        break;

                    case 6:
                        var sub6 = new CmapSubtable6Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub6);
                        subtableCache[enc.SubtableOffset] = sub6;
                        enc.Subtable = sub6;
                        break;

                    case 12:
                        var sub12 = new CmapSubtable12Deserializer(_reader).Deserialize(currentPos);
                        table.SubTables.Add(sub12);
                        subtableCache[enc.SubtableOffset] = sub12;
                        enc.Subtable = sub12;
                        break;

                    case 14:
                        // Skip format 14 (Unicode Variation Sequences)
                        var dummySubtable = new CmapSubtable14();
                        enc.IsSkipped = true;
                        enc.Subtable = dummySubtable;
                        subtableCache[enc.SubtableOffset] = dummySubtable;

                        _reader.BaseStream.Position = currentPos + 6;
                        uint length = _reader.ReadUInt32BigEndian();
                        long nextTablePos = currentPos + length;
                        if (nextTablePos > _reader.BaseStream.Length || nextTablePos < currentPos)
                        {
                            nextTablePos = _reader.BaseStream.Length;
                        }
                        _reader.BaseStream.Position = nextTablePos;
                        break;

                    default:
                        // Unsupported format - skip
                        break;
                }
            }

            return table;
        }
    }
}
