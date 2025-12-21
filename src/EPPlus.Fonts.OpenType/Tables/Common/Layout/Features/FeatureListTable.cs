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
using EPPlus.Fonts.OpenType.Subsetting;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Features
{
    /// <summary>
    /// Represents the Feature List table in GSUB, mapping features to lookup indices.
    /// </summary>
    public class FeatureListTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the list of feature records.
        /// </summary>
        public List<FeatureRecord> FeatureRecords { get; set; } = new List<FeatureRecord>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long startPos = writer.BaseStream.Position;

            // 1. Write FeatureCount
            writer.WriteUInt16BigEndian((ushort)FeatureRecords.Count);

            // 2. Write FeatureRecords (Tag + Offset)
            long offsetArrayStart = writer.BaseStream.Position;
            foreach (var record in FeatureRecords)
            {
                writer.WriteTag(record.FeatureTag);
                writer.WriteUInt16BigEndian(0); // Placeholder for offset
            }

            // 3. Serialize FeatureTables and backfill offsets
            for (int i = 0; i < FeatureRecords.Count; i++)
            {
                long currentPos = writer.BaseStream.Position;
                long recordOffsetPos = offsetArrayStart + (i * 6) + 4; // Each record is 6 bytes (4 tag + 2 offset)

                // Update the offset in the header
                this.WriteRelativeOffset(writer, startPos, recordOffsetPos);

                // Write the actual FeatureTable
                FeatureRecords[i].FeatureTable.Serialize(writer);
            }
        }

        /// <summary>
        /// Rewrites the feature list for a subset font.
        /// </summary>
        internal FeatureListTable Rewrite(FontSubsettingContext context, Dictionary<int, int> lookupMap)
        {
            System.Diagnostics.Debug.WriteLine("=== FeatureListTable.Rewrite START ===");
            System.Diagnostics.Debug.WriteLine(string.Format("Original features: {0}", this.FeatureRecords.Count));
            System.Diagnostics.Debug.WriteLine(string.Format("Lookup map has {0} entries", lookupMap != null ? lookupMap.Count : 0));

            var newList = new FeatureListTable();
            newList.FeatureRecords = new List<FeatureRecord>();

            for (int i = 0; i < this.FeatureRecords.Count; i++)
            {
                var oldRecord = this.FeatureRecords[i];

                System.Diagnostics.Debug.WriteLine(string.Format("Processing feature {0}: {1}", i, oldRecord.FeatureTag.Value));

                // Skriv om featuren med den nya lookupmappen
                var rewrittenRecord = oldRecord.Rewrite(context, lookupMap);

                // ✅ FIX: Kolla om RECORDEN är null (inte featuren)
                if (rewrittenRecord != null &&
                    rewrittenRecord.FeatureTable != null &&
                    rewrittenRecord.FeatureTable.LookupListIndices != null &&
                    rewrittenRecord.FeatureTable.LookupListIndices.Length > 0)
                {
                    newList.FeatureRecords.Add(rewrittenRecord);
                    System.Diagnostics.Debug.WriteLine(string.Format("  ✅ Kept feature with {0} lookups",
                        rewrittenRecord.FeatureTable.LookupListIndices.Length));
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("  ❌ Removed feature (no valid lookups)");
                }
            }

            System.Diagnostics.Debug.WriteLine(string.Format("=== FeatureListTable.Rewrite END: {0} features ===", newList.FeatureRecords.Count));
            return newList;
        }
    }
}