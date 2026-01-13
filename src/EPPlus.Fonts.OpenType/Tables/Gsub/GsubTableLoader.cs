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
  01/07/2026         EPPlus Software AB           Refactored to use shared loaders
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Features;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using System;
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
            _settings = tblSettings;
        }

        private readonly TableLoaderSettings _settings;

        protected override GsubTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;
            long tableStartOffset = _offset;

            // Read GSUB Header
            ushort major = _reader.ReadUInt16BigEndian();
            ushort minor = _reader.ReadUInt16BigEndian();

            if (major != 1)
            {
                // Unsupported version
                return null;
            }

            var gsubTable = new GsubTable
            {
                MajorVersion = major,
                MinorVersion = minor,
            };

            // Read Offsets
            ushort scriptListOffset = _reader.ReadUInt16BigEndian();
            ushort featureListOffset = _reader.ReadUInt16BigEndian();
            ushort lookupListOffset = _reader.ReadUInt16BigEndian();

            // ✅ Use shared loaders from Common.Layout
            if (scriptListOffset > 0)
            {
                _reader.BaseStream.Seek(tableStartOffset + scriptListOffset, SeekOrigin.Begin);
                gsubTable.ScriptList = ScriptListTableLoader.Load(_reader, tableStartOffset + scriptListOffset);
            }

            if (featureListOffset > 0)
            {
                _reader.BaseStream.Seek(tableStartOffset + featureListOffset, SeekOrigin.Begin);
                gsubTable.FeatureList = FeatureListTableLoader.Load(_reader, tableStartOffset + featureListOffset);
            }

            if (lookupListOffset > 0)
            {
                _reader.BaseStream.Seek(tableStartOffset + lookupListOffset, SeekOrigin.Begin);
                gsubTable.LookupList = LoadLookupList(tableStartOffset + lookupListOffset);
            }

            return gsubTable;
        }

        #region LookupList loading (GSUB-specific)

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

                // GSUB-specific subtable loading
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
                    default:
                        // Unknown lookup type - skip
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