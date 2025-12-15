using EPPlus.Fonts.OpenType.Tables.Gsub.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    internal class GsubTableLoader : TableLoader<GsubTable>
    {
        public GsubTableLoader(TableLoaderSettings tblSettings) : base(tblSettings, TableNames.Gsub)
        {
        }

        protected override GsubTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;
            long tableStartOffset = _offset; // GSUB tabellens startoffset

            // 1. Read GSUB Header (10 bytes)
            ushort major = _reader.ReadUInt16BigEndian();
            ushort minor = _reader.ReadUInt16BigEndian();

            if (major != 1)
            {
                // Return null if unsupported version is found
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

            // 3. Ladda ScriptList
            if (scriptListOffset > 0)
            {
                // Navigera till ScriptList start: GSUB start + ScriptList offset
                _reader.BaseStream.Seek(tableStartOffset + scriptListOffset, SeekOrigin.Begin);

                // Anropa den nya laddningsmetoden. Vi skickar med startoffseten
                // för ScriptList, eftersom ScriptTable-offsetsen är relativa till DEN.
                gsubTable.ScriptList = LoadScriptList(tableStartOffset + scriptListOffset);
            }
            // 4. Load FeatureList
            if(featureListOffset > 0)
            {
                _reader.BaseStream.Seek(tableStartOffset + featureListOffset, SeekOrigin.Begin);

                gsubTable.FeatureList = LoadFeatureList(tableStartOffset + featureListOffset);
            }

            // 5. Load LookupList
            if (lookupListOffset > 0)
            {
                // Navigate to LookupList start: GSUB start + LookupList offset
                _reader.BaseStream.Seek(tableStartOffset + lookupListOffset, SeekOrigin.Begin);

                // Anropa den nya laddningsmetoden. Skickar med LookupList startoffset.
                gsubTable.LookupList = LoadLookupList(tableStartOffset + lookupListOffset);
            }

            return gsubTable;
        }

        #region ScriptList loading

        private ScriptListTable LoadScriptList(long scriptListStartOffset)
        {
            // Positionen är redan satt till scriptListStartOffset av LoadInternal()
            ScriptListTable scriptList = new ScriptListTable();

            // USHORT ScriptCount
            ushort scriptCount = _reader.ReadUInt16BigEndian();

            // Läs in alla ScriptRecords (Tag + Offset)
            for (int i = 0; i < scriptCount; i++)
            {
                ScriptRecord record = new ScriptRecord
                {
                    // TAG ScriptTag (4 bytes)
                    ScriptTag = new Tag(_reader),
                    // USHORT ScriptOffset (relativt till scriptListStartOffset)
                    ScriptOffset = _reader.ReadUInt16BigEndian()
                };
                scriptList.ScriptRecords.Add(record);
            }

            // Spara den aktuella positionen (slutet av ScriptRecords-listan)
            // innan vi hoppar runt för att ladda ScriptTables
            long currentPosition = _reader.BaseStream.Position;

            // Ladda in ScriptTable för varje record
            foreach (var record in scriptList.ScriptRecords)
            {
                // Navigera till ScriptTable: ScriptList Start + ScriptRecord Offset
                _reader.BaseStream.Seek(scriptListStartOffset + record.ScriptOffset, SeekOrigin.Begin);

                // Ladda in ScriptTable
                record.ScriptTable = LoadScriptTable();
            }

            // Återställ läsaren till där vi var (efter ScriptRecords)
            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);

            return scriptList;
        }

        private ScriptTable LoadScriptTable()
        {
            // The reader position is already set to the ScriptTable start offset.
            long scriptTableStartOffset = _reader.BaseStream.Position;
            ScriptTable scriptTable = new ScriptTable();

            // USHORT DefaultLangSysOffset (relative to ScriptTable start position)
            ushort defaultLangSysOffset = _reader.ReadUInt16BigEndian();
            scriptTable.DefaultLangSysOffset = defaultLangSysOffset;

            // USHORT LangSysCount
            ushort langSysCount = _reader.ReadUInt16BigEndian(); // Reads into a local variable

            // Read LangSysRecord (TAG + Offset) * LangSysCount
            // We store the offsets and tags to load the tables later.
            var recordsToLoad = new Dictionary<uint, ushort>();

            for (int i = 0; i < langSysCount; i++)
            {
                // TAG LangSysTag (4 bytes)
                uint langSysTag = _reader.ReadUInt32BigEndian();

                // USHORT LangSysOffset (2 bytes, relative to ScriptTable start)
                ushort langSysOffset = _reader.ReadUInt16BigEndian();

                recordsToLoad.Add(langSysTag, langSysOffset);
            }

            // Save position after reading records to return to it later
            long positionAfterRecords = _reader.BaseStream.Position;

            // --- Deserialize LangSysTables ---
            var langSysDeserializer = new LangSysTableDeserializer(_reader);

            // 1. Load DefaultLangSysTable
            if (defaultLangSysOffset > 0)
            {
                long langSysAbsoluteStart = scriptTableStartOffset + defaultLangSysOffset;
                scriptTable.DefaultLangSys = langSysDeserializer.Deserialize(langSysAbsoluteStart);
            }

            // 2. Load other LangSysRecords
            if (langSysCount > 0)
            {
                foreach (var kvp in recordsToLoad)
                {
                    uint tag = kvp.Key;
                    ushort offset = kvp.Value;

                    long langSysAbsoluteStart = scriptTableStartOffset + offset;

                    LangSysTable langSysTable = langSysDeserializer.Deserialize(langSysAbsoluteStart);

                    scriptTable.LangSysRecords.Add(new LangSysRecord
                    {
                        LangSysTag = tag,
                        LangSysTable = langSysTable
                    });
                }
            }

            // Restore reader position to where it was after reading the record array
            _reader.BaseStream.Seek(positionAfterRecords, SeekOrigin.Begin);

            return scriptTable;
        }

        #endregion

        #region FeatureList Loading

        private FeatureListTable LoadFeatureList(long featureListStartOffset)
        {
            FeatureListTable featureList = new FeatureListTable();

            // USHORT FeatureCount
            ushort featureCount = _reader.ReadUInt16BigEndian();

            // Read all FeatureRecords (Tag + Offset)
            for (int i = 0; i < featureCount; i++)
            {
                FeatureRecord record = new FeatureRecord
                {
                    // TAG FeatureTag (4 bytes)
                    FeatureTag = new Tag(_reader),
                    // USHORT FeatureOffset (relative to featureListStartOffset)
                    FeatureOffset = _reader.ReadUInt16BigEndian()
                };
                featureList.FeatureRecords.Add(record);
            }

            // Save current position before jumping around to load FeatureTables
            long currentPosition = _reader.BaseStream.Position;

            // Load the FeatureTable for each record
            foreach (var record in featureList.FeatureRecords)
            {
                // Navigate to FeatureTable: FeatureList Start + FeatureRecord Offset
                _reader.BaseStream.Seek(featureListStartOffset + record.FeatureOffset, SeekOrigin.Begin);

                // Load the FeatureTable
                record.FeatureTable = LoadFeatureTable();
            }

            // Restore reader position
            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);

            return featureList;
        }

        private FeatureTable LoadFeatureTable()
        {
            FeatureTable featureTable = new FeatureTable();

            // USHORT FeatureParams (reserved for future use, should be 0)
            featureTable.FeatureParams = _reader.ReadUInt16BigEndian();

            // USHORT LookupCount
            featureTable.LookupCount = _reader.ReadUInt16BigEndian();

            // USHORT[] LookupListIndices: Read all indices
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
                // .NET 3.5 compatible replacement for Array.Empty<ushort>()
                featureTable.LookupListIndices = new ushort[0];
            }

            return featureTable;
        }

        #endregion

        #region LookupList loading
        private LookupListTable LoadLookupList(long lookupListStartOffset)
        {
            LookupListTable lookupList = new LookupListTable();

            // USHORT LookupCount
            ushort lookupCount = _reader.ReadUInt16BigEndian();

            // USHORT[] LookupOffsets (relative to lookupListStartOffset)
            ushort[] lookupOffsets = new ushort[lookupCount];
            for (int i = 0; i < lookupCount; i++)
            {
                lookupOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            // Save current position before jumping around to load LookupTables
            long currentPosition = _reader.BaseStream.Position;

            // Load the LookupTable for each offset
            foreach (ushort offset in lookupOffsets)
            {
                // Navigate to LookupTable: LookupList Start + Lookup Offset
                _reader.BaseStream.Seek(lookupListStartOffset + offset, SeekOrigin.Begin);

                // Load the LookupTable
                LookupTable lookupTable = LoadLookupTable();
                lookupList.Lookups.Add(lookupTable);
            }

            // Restore reader position
            _reader.BaseStream.Seek(currentPosition, SeekOrigin.Begin);

            return lookupList;
        }

        private LookupTable LoadLookupTable()
        {
            // The reader position is already set to the LookupTable start offset by the calling method.
            LookupTable lookupTable = new LookupTable();

            // Store the start offset for calculating relative SubTable offsets later
            long lookupTableStartOffset = _reader.BaseStream.Position;

            // USHORT LookupType (1=Single, 4=Ligature, etc.)
            lookupTable.LookupType = _reader.ReadUInt16BigEndian();

            // USHORT LookupFlag
            lookupTable.LookupFlag = _reader.ReadUInt16BigEndian();

            // USHORT SubTableCount
            lookupTable.SubTableCount = _reader.ReadUInt16BigEndian();

            // Read USHORT[] SubTableOffsets (relative to this LookupTable start)
            ushort[] subTableOffsets = new ushort[lookupTable.SubTableCount];
            for (int i = 0; i < lookupTable.SubTableCount; i++)
            {
                subTableOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            // Save the current position after reading all offsets, before diving into SubTables
            long positionAfterOffsets = _reader.BaseStream.Position;

            // --- Deserialize SubTables based on LookupType ---
            for (int i = 0; i < lookupTable.SubTableCount; i++)
            {
                // Calculate the absolute start offset of the SubTable
                long subTableAbsoluteStart = lookupTableStartOffset + subTableOffsets[i];

                // Determine which deserializer to use
                FontTableElement subTable = null;

                switch (lookupTable.LookupType)
                {
                    case 1: // Single Substitution
                        var singleLoader = new SingleSubstSubTableDeserializer(_reader);
                        subTable = singleLoader.Deserialize(subTableAbsoluteStart);
                        break;

                    case 4: // Ligature Substitution
                        var ligLoader = new LigatureSubstSubTableDeserializer(_reader);
                        subTable = ligLoader.Deserialize(subTableAbsoluteStart);
                        break;

                    default:
                        // Unsupported LookupType. Skip loading for this subtable.
                        break;
                }

                if (subTable != null)
                {
                    lookupTable.SubTables.Add(subTable);
                }
            }

            // Restore the reader position to where it was after reading the offset array, 
            // before the subtable deserialization loop started.
            _reader.BaseStream.Seek(positionAfterOffsets, SeekOrigin.Begin);

            return lookupTable;
        }
        #endregion
    }
}
