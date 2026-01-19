/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           Shared FeatureList loader
 *************************************************************************************************/
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Features
{
    /// <summary>
    /// Shared loader for FeatureListTable used by both GSUB and GPOS
    /// </summary>
    internal static class FeatureListTableLoader
    {
        public static FeatureListTable Load(FontsBinaryReader reader, long featureListStart)
        {
            var featureList = new FeatureListTable();
            ushort featureCount = reader.ReadUInt16BigEndian();

            var featureOffsets = new List<FeatureOffsetRecord>();
            for (int i = 0; i < featureCount; i++)
            {
                var tag = new Tag(reader);
                ushort offset = reader.ReadUInt16BigEndian();
                featureOffsets.Add(new FeatureOffsetRecord { Tag = tag, Offset = offset });
            }

            long positionAfterRecords = reader.BaseStream.Position;

            // Load feature tables
            foreach (var record in featureOffsets)
            {
                reader.BaseStream.Seek(featureListStart + record.Offset, SeekOrigin.Begin);
                var featureTable = LoadFeatureTable(reader);

                featureList.FeatureRecords.Add(new FeatureRecord
                {
                    FeatureTag = record.Tag,
                    FeatureOffset = record.Offset,
                    FeatureTable = featureTable
                });
            }

            reader.BaseStream.Seek(positionAfterRecords, SeekOrigin.Begin);
            return featureList;
        }

        private static FeatureTable LoadFeatureTable(FontsBinaryReader reader)
        {
            var featureTable = new FeatureTable();
            featureTable.FeatureParams = reader.ReadUInt16BigEndian();
            featureTable.LookupCount = reader.ReadUInt16BigEndian();

            if (featureTable.LookupCount > 0)
            {
                featureTable.LookupListIndices = new ushort[featureTable.LookupCount];
                for (int i = 0; i < featureTable.LookupCount; i++)
                {
                    featureTable.LookupListIndices[i] = reader.ReadUInt16BigEndian();
                }
            }
            else
            {
                featureTable.LookupListIndices = new ushort[0];
            }

            return featureTable;
        }

        // Helper struct for .NET 3.5 compatibility (no tuples)
        private struct FeatureOffsetRecord
        {
            public Tag Tag;
            public ushort Offset;
        }
    }
}