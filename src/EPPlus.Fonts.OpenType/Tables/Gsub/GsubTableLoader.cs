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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    /// <summary>
    /// Loads the GSUB (Glyph Substitution) table from an OpenType font.
    /// Manages the hierarchical loading of Scripts, Features, and Lookups.
    /// </summary>
    internal class GsubTableLoader : TableLoader<GsubTable>
    {
        public GsubTableLoader(TableLoaderSettings tblSettings) : base(tblSettings, TableNames.Gsub)
        {
        }

        protected override GsubTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;
            long tableStartOffset = _offset; // The start of the GSUB table in the stream

            // 1. Read GSUB Header (10 bytes for version 1.0)
            ushort major = _reader.ReadUInt16BigEndian();
            ushort minor = _reader.ReadUInt16BigEndian();

            if (major != 1)
            {
                // Unsupported version found
                return null;
            }

            var gsubTable = new GsubTable
            {
                MajorVersion = major,
                MinorVersion = minor,
            };

            // 2. Read Offsets (USHORT, relative to tableStartOffset)
            ushort scriptListOffset = _reader.ReadUInt16BigEndian();
            ushort featureListOffset = _reader.ReadUInt16BigEndian();
            ushort lookupListOffset = _reader.ReadUInt16BigEndian();

            // 3. Load ScriptList
            if (scriptListOffset > 0)
            {
                // Navigate to ScriptList start: GSUB start + ScriptList offset
                _reader.BaseStream.Seek(tableStartOffset + scriptListOffset, SeekOrigin.Begin);

                // Load the ScriptList; internal offsets within it are relative to its own start
                gsubTable.ScriptList = LoadScriptList(tableStartOffset + scriptListOffset);
            }

            // 4. Load FeatureList
            if (featureListOffset > 0)
            {
                _reader.BaseStream.Seek(tableStartOffset + featureListOffset, SeekOrigin.Begin);
                gsubTable.FeatureList = LoadFeatureList(tableStartOffset + featureListOffset);
            }

            // 5. Load LookupList
            if (lookupListOffset > 0)
            {
                _reader.BaseStream.Seek(tableStartOffset + lookupListOffset, SeekOrigin.Begin);
                gsubTable.LookupList = LoadLookupList(tableStartOffset + lookupListOffset);
            }

            return gsubTable;
        }

        #region ScriptList loading

        private ScriptListTable LoadScriptList(long scriptListStartOffset)
        {
            ScriptListTable scriptList = new ScriptListTable();

            // USHORT ScriptCount
            ushort scriptCount = _reader.ReadUInt16BigEndian();

            // Read all ScriptRecords (Tag + Offset)
            for (int i = 0; i < scriptCount; i++)
            {
                ScriptRecord record = new ScriptRecord
                {
                    ScriptTag = new Tag(_reader),
                    // Offset relative to scriptListStartOffset
                    ScriptOffset = _reader.ReadUInt16BigEndian()
                };
                scriptList.ScriptRecords.Add(record);
            }

            // Save the position at the end of ScriptRecords before jumping to ScriptTables
            long currentPosition = _reader.BaseStream.Position;

            // Load ScriptTable for each record
            foreach (var record in scriptList.ScriptRecords)
            {
                _reader.BaseStream.Seek(scriptListStartOffset + record.ScriptOffset, SeekOrigin.Begin);
                record.ScriptTable = LoadScriptTable();
            }

            // Restore position to continue sequential reading if necessary
            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);

            return scriptList;
        }

        private ScriptTable LoadScriptTable()
        {
            long scriptTableStartOffset = _reader.BaseStream.Position;
            ScriptTable scriptTable = new ScriptTable();

            // USHORT DefaultLangSysOffset (relative to ScriptTable start)
            ushort defaultLangSysOffset = _reader.ReadUInt16BigEndian();
            scriptTable.DefaultLangSysOffset = defaultLangSysOffset;

            // USHORT LangSysCount
            ushort langSysCount = _reader.ReadUInt16BigEndian();

            // Store records to load after the record array is fully read
            var recordsToLoad = new Dictionary<uint, ushort>();

            for (int i = 0; i < langSysCount; i++)
            {
                uint langSysTag = _reader.ReadUInt32BigEndian();
                ushort langSysOffset = _reader.ReadUInt16BigEndian();
                recordsToLoad.Add(langSysTag, langSysOffset);
            }

            long positionAfterRecords = _reader.BaseStream.Position;
            var langSysDeserializer = new LangSysTableDeserializer(_reader);

            // 1. Load DefaultLangSysTable
            if (defaultLangSysOffset > 0)
            {
                long langSysAbsoluteStart = scriptTableStartOffset + defaultLangSysOffset;
                scriptTable.DefaultLangSys = langSysDeserializer.Deserialize(langSysAbsoluteStart);
            }

            // 2. Load other LangSysRecords
            foreach (var kvp in recordsToLoad)
            {
                long langSysAbsoluteStart = scriptTableStartOffset + kvp.Value;
                scriptTable.LangSysRecords.Add(new LangSysRecord
                {
                    LangSysTag = kvp.Key,
                    LangSysTable = langSysDeserializer.Deserialize(langSysAbsoluteStart)
                });
            }

            _reader.BaseStream.Seek(positionAfterRecords, SeekOrigin.Begin);
            return scriptTable;
        }

        #endregion

        #region FeatureList Loading

        private FeatureListTable LoadFeatureList(long featureListStartOffset)
        {
            FeatureListTable featureList = new FeatureListTable();
            ushort featureCount = _reader.ReadUInt16BigEndian();

            for (int i = 0; i < featureCount; i++)
            {
                featureList.FeatureRecords.Add(new FeatureRecord
                {
                    FeatureTag = new Tag(_reader),
                    FeatureOffset = _reader.ReadUInt16BigEndian()
                });
            }

            long currentPosition = _reader.BaseStream.Position;

            foreach (var record in featureList.FeatureRecords)
            {
                _reader.BaseStream.Seek(featureListStartOffset + record.FeatureOffset, SeekOrigin.Begin);
                record.FeatureTable = LoadFeatureTable();
            }

            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);
            return featureList;
        }

        private FeatureTable LoadFeatureTable()
        {
            FeatureTable featureTable = new FeatureTable();
            featureTable.FeatureParams = _reader.ReadUInt16BigEndian();
            featureTable.LookupCount = _reader.ReadUInt16BigEndian();

            if (featureTable.LookupCount > 0)
            {
                featureTable.LookupListIndices = new ushort[featureTable.LookupCount];
                for (int i = 0; i < featureTable.LookupCount; i++)
                {
                    featureTable.LookupListIndices[i] = _reader.ReadUInt16BigEndian();
                }
            }
            else
            {
                featureTable.LookupListIndices = new ushort[0];
            }

            return featureTable;
        }

        #endregion

        #region LookupList loading

        private LookupListTable LoadLookupList(long lookupListStartOffset)
        {
            LookupListTable lookupList = new LookupListTable();
            ushort lookupCount = _reader.ReadUInt16BigEndian();

            ushort[] lookupOffsets = new ushort[lookupCount];
            for (int i = 0; i < lookupCount; i++)
            {
                lookupOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            long currentPosition = _reader.BaseStream.Position;

            foreach (ushort offset in lookupOffsets)
            {
                _reader.BaseStream.Seek(lookupListStartOffset + offset, SeekOrigin.Begin);
                lookupList.Lookups.Add(LoadLookupTable());
            }

            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);
            return lookupList;
        }

        private LookupTable LoadLookupTable()
        {
            LookupTable lookupTable = new LookupTable();
            long lookupTableStartOffset = _reader.BaseStream.Position;

            lookupTable.LookupType = _reader.ReadUInt16BigEndian();
            lookupTable.LookupFlag = _reader.ReadUInt16BigEndian();
            lookupTable.SubTableCount = _reader.ReadUInt16BigEndian();

            ushort[] subTableOffsets = new ushort[lookupTable.SubTableCount];
            for (int i = 0; i < lookupTable.SubTableCount; i++)
            {
                subTableOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            long positionAfterOffsets = _reader.BaseStream.Position;

            for (int i = 0; i < lookupTable.SubTableCount; i++)
            {
                long subTableAbsoluteStart = lookupTableStartOffset + subTableOffsets[i];
                FontTableElement subTable = null;

                switch (lookupTable.LookupType)
                {
                    case 1: // Single Substitution
                        subTable = new SingleSubstSubTableDeserializer(_reader).Deserialize(subTableAbsoluteStart);
                        break;
                    case 4: // Ligature Substitution
                        subTable = new LigatureSubstSubTableDeserializer(_reader).Deserialize(subTableAbsoluteStart);
                        break;
                    case 6: // Chaining Contextual Substitution
                        subTable = new ChainingContextualDeserializer(_reader).Deserialize(subTableAbsoluteStart);
                        break;
                    case 7: // Extension Substitution
                        subTable = new ExtensionSubstSubTableDeserializer(_reader).Deserialize(subTableAbsoluteStart);
                        break;
                }

                if (subTable != null)
                {
                    lookupTable.SubTables.Add(subTable);
                }
            }

            _reader.BaseStream.Seek(positionAfterOffsets, SeekOrigin.Begin);
            return lookupTable;
        }
        #endregion
    }
}