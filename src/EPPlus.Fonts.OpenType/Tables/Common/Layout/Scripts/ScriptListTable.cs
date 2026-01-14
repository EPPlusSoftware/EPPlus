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
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts
{
    /// <summary>
    /// Represents the Script List table in an OpenType font.
    /// It identifies the scripts in the font and points to the Script tables that define language systems.
    /// </summary>
    public class ScriptListTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the list of Script Records.
        /// </summary>
        public List<ScriptRecord> ScriptRecords { get; set; } = new List<ScriptRecord>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Store the start position of the ScriptList for relative offset calculations
            long scriptListStartOffset = writer.BaseStream.Position;

            // 1. Write ScriptCount (USHORT)
            writer.WriteUInt16BigEndian((ushort)this.ScriptRecords.Count);

            // 2. Write ScriptRecords (Tag + Placeholder for Offset)
            List<long> recordOffsetPositions = new List<long>();

            foreach (var record in this.ScriptRecords)
            {
                // Write ScriptTag (4 bytes)
                writer.Write(record.ScriptTag.ToBytes());

                // Store position for the ScriptTableOffset (2 bytes) placeholder
                recordOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- Serialize ScriptTables ---

            // We iterate through the records and serialize the associated ScriptTable for each.
            for (int i = 0; i < this.ScriptRecords.Count; i++)
            {
                var record = this.ScriptRecords[i];
                long currentOffset = writer.BaseStream.Position;

                // Calculate the offset relative to the start of the ScriptListTable
                ushort relativeScriptTableOffset = (ushort)(currentOffset - scriptListStartOffset);

                // 1. Return to the record's offset field and backfill the calculated value
                long recordOffsetPos = recordOffsetPositions[i];
                writer.BaseStream.Seek(recordOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeScriptTableOffset);

                // 2. Restore the stream position to continue serializing data
                writer.BaseStream.Seek(currentOffset, SeekOrigin.Begin);

                // Serialize the ScriptTable itself (must implement Serialize)
                if (record.ScriptTable != null)
                {
                    record.ScriptTable.Serialize(writer);
                }
            }
        }

        internal ScriptListTable Rewrite(FontSubsettingContext context, Dictionary<int, int> featureIndexMap)
        {
            var newScriptList = new ScriptListTable();

            foreach (var scriptRecord in this.ScriptRecords)
            {
                var newScriptRecord = new ScriptRecord
                {
                    ScriptTag = scriptRecord.ScriptTag,
                    ScriptTable = RewriteScriptTable(scriptRecord.ScriptTable, featureIndexMap)
                };

                newScriptList.ScriptRecords.Add(newScriptRecord);
            }

            return newScriptList;
        }

        /// <summary>
        /// Rewrites a ScriptTable by remapping feature indices.
        /// </summary>
        private ScriptTable RewriteScriptTable(ScriptTable original, Dictionary<int, int> featureIndexMap)
        {
            var newScriptTable = new ScriptTable
            {
                DefaultLangSysOffset = original.DefaultLangSysOffset
            };

            // Rewrite DefaultLangSys
            if (original.DefaultLangSys != null)
            {
                newScriptTable.DefaultLangSys = RewriteLangSys(original.DefaultLangSys, featureIndexMap);
            }

            // Rewrite LangSysRecords
            foreach (var langSysRecord in original.LangSysRecords)
            {
                var newLangSys = RewriteLangSys(langSysRecord.LangSysTable, featureIndexMap);
                if (newLangSys != null && newLangSys.FeatureIndices.Length > 0)
                {
                    newScriptTable.LangSysRecords.Add(new LangSysRecord
                    {
                        LangSysTag = langSysRecord.LangSysTag,
                        LangSysTable = newLangSys
                    });
                }
            }

            return newScriptTable;
        }

        /// <summary>
        /// Rewrites a LangSysTable by remapping feature indices.
        /// </summary>
        private LangSysTable RewriteLangSys(LangSysTable original, Dictionary<int, int> featureIndexMap)
        {
            if (original == null)
                return null;

            var newFeatureIndices = new List<ushort>();

            // Remap each feature index
            foreach (var oldIndex in original.FeatureIndices)
            {
                if (featureIndexMap.TryGetValue(oldIndex, out int newIndex))
                {
                    newFeatureIndices.Add((ushort)newIndex);
                }
                // Else: feature was removed, skip it
            }

            // Handle RequiredFeatureIndex
            ushort newRequiredFeatureIndex = 0xFFFF; // Default: no required feature
            if (original.RequiredFeatureIndex != 0xFFFF)
            {
                if (featureIndexMap.TryGetValue(original.RequiredFeatureIndex, out int mappedRequired))
                {
                    newRequiredFeatureIndex = (ushort)mappedRequired;
                }
                // Else: required feature was removed, set to 0xFFFF
            }

            return new LangSysTable
            {
                LookupOrder = original.LookupOrder,
                RequiredFeatureIndex = newRequiredFeatureIndex,
                FeatureIndexCount = (ushort)newFeatureIndices.Count,
                FeatureIndices = newFeatureIndices.ToArray()
            };
        }
    }
}