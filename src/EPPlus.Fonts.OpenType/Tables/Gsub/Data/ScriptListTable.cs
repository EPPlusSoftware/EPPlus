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

        internal ScriptListTable Rewrite(EPPlus.Fonts.OpenType.Subsetting.FontSubsettingContext context)
        {
            var newList = new ScriptListTable();

            foreach (var oldRecord in this.ScriptRecords)
            {
                var newRecord = new ScriptRecord();
                newRecord.ScriptTag = oldRecord.ScriptTag;

                if (oldRecord.ScriptTable != null)
                {
                    var oldTable = oldRecord.ScriptTable;
                    var newTable = new ScriptTable();

                    // Copy DefaultLangSys
                    if (oldTable.DefaultLangSys != null)
                    {
                        var oldLang = oldTable.DefaultLangSys;
                        var newLang = new LangSysTable();
                        newLang.LookupOrder = oldLang.LookupOrder;
                        newLang.RequiredFeatureIndex = oldLang.RequiredFeatureIndex;
                        newLang.FeatureIndexCount = oldLang.FeatureIndexCount;

                        if (oldLang.FeatureIndices != null)
                        {
                            var newIndices = new ushort[oldLang.FeatureIndices.Length];
                            for (int i = 0; i < oldLang.FeatureIndices.Length; i++)
                            {
                                newIndices[i] = oldLang.FeatureIndices[i];
                            }
                            newLang.FeatureIndices = newIndices;
                        }
                        newTable.DefaultLangSys = newLang;
                    }

                    // Copy LangSysRecords
                    if (oldTable.LangSysRecords != null)
                    {
                        newTable.LangSysRecords = new List<LangSysRecord>();
                        foreach (var oldLangRecord in oldTable.LangSysRecords)
                        {
                            var newLangRecord = new LangSysRecord();
                            newLangRecord.LangSysTag = oldLangRecord.LangSysTag;

                            if (oldLangRecord.LangSysTable != null)
                            {
                                var oldL = oldLangRecord.LangSysTable;
                                var newL = new LangSysTable();
                                newL.LookupOrder = oldL.LookupOrder;
                                newL.RequiredFeatureIndex = oldL.RequiredFeatureIndex;
                                newL.FeatureIndexCount = oldL.FeatureIndexCount;

                                if (oldL.FeatureIndices != null)
                                {
                                    var newIndices = new ushort[oldL.FeatureIndices.Length];
                                    for (int j = 0; j < oldL.FeatureIndices.Length; j++)
                                    {
                                        newIndices[j] = oldL.FeatureIndices[j];
                                    }
                                    newL.FeatureIndices = newIndices;
                                }
                                newLangRecord.LangSysTable = newL;
                            }
                            newTable.LangSysRecords.Add(newLangRecord);
                        }
                    }

                    newRecord.ScriptTable = newTable;
                }

                newList.ScriptRecords.Add(newRecord);
            }

            return newList;
        }
    }
}