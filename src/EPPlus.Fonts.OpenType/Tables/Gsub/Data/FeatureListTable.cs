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
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    /// <summary>
    /// Represents the Feature List Table in an OpenType font, containing records of font features.
    /// </summary>
    public class FeatureListTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the list of feature records.
        /// </summary>
        public List<FeatureRecord> FeatureRecords { get; set; } = new List<FeatureRecord>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Write the number of feature records
            writer.WriteUInt16BigEndian((ushort)this.FeatureRecords.Count);

            // Track positions for backfilling offsets
            List<long> recordOffsetPositions = new List<long>();

            // Offsets in FeatureList are relative to the start of the FeatureList table
            long featureListStartOffset = writer.BaseStream.Position - sizeof(ushort);

            foreach (var record in this.FeatureRecords)
            {
                // Write FeatureTag (4 bytes)
                writer.Write(record.FeatureTag.ToBytes());

                // Store position for the FeatureTableOffset (2 bytes) to be backfilled later
                recordOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- Serialize FeatureTables ---

            int recordIndex = 0;
            foreach (var record in this.FeatureRecords)
            {
                long currentTablePosition = writer.BaseStream.Position;

                // Calculate the offset relative to the start of the FeatureList
                ushort relativeFeatureTableOffset = (ushort)(currentTablePosition - featureListStartOffset);

                // 1. Backfill the offset in the corresponding FeatureRecord
                long recordOffsetPos = recordOffsetPositions[recordIndex];
                writer.BaseStream.Seek(recordOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeFeatureTableOffset);

                // 2. Return to the end of the stream and serialize the actual FeatureTable
                writer.BaseStream.Seek(currentTablePosition, SeekOrigin.Begin);
                record.FeatureTable.Serialize(writer);

                recordIndex++;
            }
        }
    }
}