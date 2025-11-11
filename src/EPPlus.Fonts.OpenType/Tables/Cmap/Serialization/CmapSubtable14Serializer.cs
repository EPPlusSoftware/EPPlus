using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable14Serializer : CmapSubtableSerializerBase<CmapSubtable14>
    {
        internal override void Serialize(CmapSubtable14 subTable, FontsBinaryWriter writer)
        {
            // Save the starting position of the subtable in the stream
            long startOffset = writer.BaseStream.Position;

            // Write the header:
            // Format (2 bytes), Length (4 bytes placeholder), Number of VariationSelector records (4 bytes)
            writer.WriteUInt16BigEndian(subTable.Format);       // Format = 14
            writer.WriteUInt32BigEndian(0);                     // Placeholder for Length
            writer.WriteUInt32BigEndian((uint)subTable.VariationSelectors.Count);

            // Store the positions where we will later write the DefaultUVSOffset and NonDefaultUVSOffset
            var selectorOffsetPositions = new List<long>();

            // Write each VariationSelector record with placeholder offsets
            foreach (var selector in subTable.VariationSelectors)
            {
                writer.WriteUInt24BigEndian(selector.VarSelector);

                // Save the position where the offsets will be written
                selectorOffsetPositions.Add(writer.BaseStream.Position);

                writer.WriteUInt32BigEndian(0); // DefaultUVSOffset placeholder
                writer.WriteUInt32BigEndian(0); // NonDefaultUVSOffset placeholder
            }

            // Write UVS tables and record their offsets
            for (int i = 0; i < subTable.VariationSelectors.Count; i++)
            {
                var selector = subTable.VariationSelectors[i];

                // Write Default UVS Table if present
                if (selector.DefaultUvsTable != null && selector.DefaultUvsTable.Ranges.Count > 0)
                {
                    long offset = writer.BaseStream.Position - startOffset;
                    selector.DefaultUVSOffset = (uint)offset;
                    selector.DefaultUvsTable.Serialize(writer);
                }

                // Write Non-Default UVS Table if present
                if (selector.NonDefaultUvsTable != null && selector.NonDefaultUvsTable.Mappings.Count > 0)
                {
                    long offset = writer.BaseStream.Position - startOffset;
                    selector.NonDefaultUVSOffset = (uint)offset;
                    selector.NonDefaultUvsTable.Serialize(writer);
                }

                // Go back and write the correct offsets into the selector record
                long offsetPos = selectorOffsetPositions[i];
                long currentPos = writer.BaseStream.Position;

                writer.BaseStream.Position = offsetPos;
                writer.WriteUInt32BigEndian(selector.DefaultUVSOffset);
                writer.WriteUInt32BigEndian(selector.NonDefaultUVSOffset);

                // Return to the current position to continue writing
                writer.BaseStream.Position = currentPos;
            }

            // Calculate and update the total length of the subtable
            long endOffset = writer.BaseStream.Position;
            subTable.Length = (uint)(endOffset - startOffset);

            // Go back and write the correct length into the header
            long lengthPos = startOffset + 2; // Length field starts after Format (2 bytes)
            long finalPos = writer.BaseStream.Position;

            writer.BaseStream.Position = lengthPos;
            writer.WriteUInt32BigEndian(subTable.Length);

            // Return to the end of the stream
            writer.BaseStream.Position = finalPos;
        }
    }
}
