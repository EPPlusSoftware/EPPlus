/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS table loader
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.IO;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos
{
    /// <summary>
    /// Loader for GPOS (Glyph Positioning) table
    /// </summary>
    internal class GposTableLoader : TableLoader<GposTable>
    {
        public GposTableLoader(TableLoaderSettings settings)
            : base(settings, TableNames.Gpos)
        {
        }

        protected override GposTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;
            var table = new GposTable();

            // Read version
            table.MajorVersion = _reader.ReadUInt16BigEndian();
            table.MinorVersion = _reader.ReadUInt16BigEndian();

            if (table.MajorVersion != 1)
            {
                throw new NotSupportedException($"GPOS version {table.MajorVersion}.{table.MinorVersion} is not supported. Only version 1.x is supported.");
            }

            // Read offsets
            ushort scriptListOffset = _reader.ReadUInt16BigEndian();
            ushort featureListOffset = _reader.ReadUInt16BigEndian();
            ushort lookupListOffset = _reader.ReadUInt16BigEndian();

            uint featureVariationsOffset = 0;
            if (table.MinorVersion >= 1)
            {
                featureVariationsOffset = _reader.ReadUInt32BigEndian();
                table.FeatureVariationsOffset = featureVariationsOffset;
            }

            long tableStart = _offset;

            // ✅ Use shared loaders from Common.Layout
            if (scriptListOffset > 0)
            {
                _reader.BaseStream.Position = tableStart + scriptListOffset;
                table.ScriptList = ScriptListTableLoader.Load(_reader, tableStart + scriptListOffset);
            }

            if (featureListOffset > 0)
            {
                _reader.BaseStream.Position = tableStart + featureListOffset;
                table.FeatureList = FeatureListTableLoader.Load(_reader, tableStart + featureListOffset);
            }

            if (lookupListOffset > 0)
            {
                _reader.BaseStream.Position = tableStart + lookupListOffset;
                table.LookupList = ReadLookupList(_reader, tableStart + lookupListOffset);
            }

            // Note: FeatureVariations not implemented yet
            // When implemented, read at tableStart + featureVariationsOffset

            return table;
        }

        private LookupListTable ReadLookupList(FontsBinaryReader reader, long lookupListStart)
        {
            var lookupList = new LookupListTable();

            ushort lookupCount = reader.ReadUInt16BigEndian();
            var lookupOffsets = new ushort[lookupCount];

            for (int i = 0; i < lookupCount; i++)
            {
                lookupOffsets[i] = reader.ReadUInt16BigEndian();
            }

            lookupList.Lookups = new List<LookupTable>();

            for (int i = 0; i < lookupCount; i++)
            {
                if (lookupOffsets[i] == 0)
                    continue;

                reader.BaseStream.Position = lookupListStart + lookupOffsets[i];
                var lookup = ReadLookup(reader, lookupListStart + lookupOffsets[i]);
                lookupList.Lookups.Add(lookup);
            }

            return lookupList;
        }

        private LookupTable ReadLookup(FontsBinaryReader reader, long lookupStart)
        {
            var lookup = new LookupTable();

            lookup.LookupType = reader.ReadUInt16BigEndian();
            lookup.LookupFlag = reader.ReadUInt16BigEndian();
            ushort subTableCount = reader.ReadUInt16BigEndian();

            var subTableOffsets = new ushort[subTableCount];
            for (int i = 0; i < subTableCount; i++)
            {
                subTableOffsets[i] = reader.ReadUInt16BigEndian();
            }

            // Read MarkFilteringSet if present
            if ((lookup.LookupFlag & 0x0010) != 0)
            {
                lookup.MarkFilteringSet = reader.ReadUInt16BigEndian();
            }

            lookup.SubTables = new List<FontTableElement>();

            for (int i = 0; i < subTableCount; i++)
            {
                if (subTableOffsets[i] == 0)
                    continue;

                reader.BaseStream.Position = lookupStart + subTableOffsets[i];
                var subTable = ReadSubTable(reader, lookup.LookupType, lookupStart + subTableOffsets[i]);

                if (subTable != null)
                {
                    lookup.SubTables.Add(subTable);
                }
            }

            return lookup;
        }

        private FontTableElement ReadSubTable(FontsBinaryReader reader, ushort lookupType, long subtableStart)
        {
            switch (lookupType)
            {
                case 1:
                    // Single adjustment positioning
                    // TODO: Implement when needed
                    return null;

                case 2:
                    // Pair adjustment positioning (KERNING!)
                    return ReadPairPosSubTable(reader, subtableStart);

                case 3:
                    // Cursive attachment positioning
                    // TODO: Implement when needed
                    return null;

                case 4:
                    // MarkToBase attachment positioning
                    // TODO: Implement when needed
                    return null;

                case 5:
                    // MarkToLigature attachment positioning
                    // TODO: Implement when needed
                    return null;

                case 6:
                    // MarkToMark attachment positioning
                    // TODO: Implement when needed
                    return null;

                case 7:
                    // Context positioning
                    // TODO: Implement when needed
                    return null;

                case 8:
                    // Chained context positioning
                    // TODO: Implement when needed
                    return null;

                case 9:
                    // Extension positioning
                    // TODO: Implement when needed
                    return null;

                default:
                    // Unknown lookup type - skip
                    return null;
            }
        }

        private FontTableElement ReadPairPosSubTable(FontsBinaryReader reader, long subtableStart)
        {
            ushort posFormat = reader.ReadUInt16BigEndian();

            if (posFormat == 1)
            {
                return ReadPairPosFormat1(reader, subtableStart);
            }
            else if (posFormat == 2)
            {
                // TODO: Implement Format 2 (class-based pairs)
                return null;
            }
            else
            {
                throw new NotSupportedException($"PairPos format {posFormat} is not supported.");
            }
        }

        private PairPosSubTableFormat1 ReadPairPosFormat1(FontsBinaryReader reader, long subtableStart)
        {
            // ✅ Use the deserializer!
            var deserializer = new PairPosSubTableFormat1Deserializer(reader);
            return deserializer.Deserialize(subtableStart);
        }
    }
}