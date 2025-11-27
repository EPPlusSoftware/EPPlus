using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable14Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CmapSubtable14Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        internal CmapSubtable14 Deserialize(uint startIndex)
        {
            // Move to the start of the subtable
            _reader.BaseStream.Position = startIndex;
            long startOffset = _reader.BaseStream.Position;

            // Read header
            ushort format = _reader.ReadUInt16BigEndian(); // Should be 14
            if (format != 14)
                throw new InvalidDataException("Expected format 14");

            uint length = _reader.ReadUInt32BigEndian();
            uint numVarSelectorRecords = _reader.ReadUInt32BigEndian();

            var subtable = new CmapSubtable14
            {
                Length = length,
                Language = 0 // Format 14 does not use language
            };

            // Read all VariationSelector records
            for (int i = 0; i < numVarSelectorRecords; i++)
            {
                uint varSelector = _reader.ReadUInt24BigEndian();
                uint defaultUVSOffset = _reader.ReadUInt32BigEndian();
                uint nonDefaultUVSOffset = _reader.ReadUInt32BigEndian();

                var selector = new VariationSelector
                {
                    VarSelector = varSelector,
                    DefaultUVSOffset = defaultUVSOffset,
                    NonDefaultUVSOffset = nonDefaultUVSOffset
                };

                // Read Default UVS Table if present
                if (defaultUVSOffset != 0)
                {
                    _reader.BaseStream.Position = startOffset + defaultUVSOffset;
                    uint numRanges = _reader.ReadUInt32BigEndian();

                    var defaultTable = new DefaultUvsTable();
                    for (int j = 0; j < numRanges; j++)
                    {
                        var range = new UnicodeRange
                        {
                            StartUnicodeValue = _reader.ReadUInt24BigEndian(),
                            AdditionalCount = _reader.ReadByte()
                        };
                        defaultTable.Ranges.Add(range);
                    }

                    selector.DefaultUvsTable = defaultTable;
                }

                // Read Non-Default UVS Table if present
                if (nonDefaultUVSOffset != 0)
                {
                    _reader.BaseStream.Position = startOffset + nonDefaultUVSOffset;
                    uint numMappings = _reader.ReadUInt32BigEndian();

                    var nonDefaultTable = new NonDefaultUvsTable();
                    for (int j = 0; j < numMappings; j++)
                    {
                        var mapping = new UvsMapping
                        {
                            UnicodeValue = _reader.ReadUInt24BigEndian(),
                            GlyphId = _reader.ReadUInt16BigEndian()
                        };
                        nonDefaultTable.Mappings.Add(mapping);
                    }

                    selector.NonDefaultUvsTable = nonDefaultTable;
                }

                subtable.VariationSelectors.Add(selector);
            }

            return subtable;
        }
    }
}
