using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LookupTable : FontTableElement
    {
        public ushort LookupType { get; set; }

        public ushort LookupFlag { get; set; }

        public ushort SubTableCount { get; set; }

        public List<FontTableElement> SubTables { get; set; } = new List<FontTableElement>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // USHORT LookupType
            writer.WriteUInt16BigEndian(this.LookupType);

            // USHORT LookupFlag
            writer.WriteUInt16BigEndian(this.LookupFlag);

            // USHORT SubTableCount
            writer.WriteUInt16BigEndian((ushort)this.SubTables.Count);

            // Placeholder for USHORT[] SubTableOffsets (relative to this LookupTable start)
            List<long> subTableOffsetPositions = new List<long>();
            for (int i = 0; i < this.SubTables.Count; i++)
            {
                subTableOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0); // Placeholder
            }

            // Store the start of the LookupTable for offset calculation
            long lookupTableStartOffset = writer.BaseStream.Position - (sizeof(ushort) * this.SubTables.Count) - (sizeof(ushort) * 3);

            // --- Skriv ut Sub-tabellerna ---

            int subIndex = 0;
            foreach (FontTableElement subTable in this.SubTables)
            {
                long currentOffset = writer.BaseStream.Position;

                // Skriv ut den relativa offseten till SubTable
                long subTableOffsetPos = subTableOffsetPositions[subIndex];
                // Offset is relative to the start of the LookupTable
                ushort relativeSubTableOffset = (ushort)(currentOffset - lookupTableStartOffset);

                // Gå tillbaka och skriv offseten
                writer.BaseStream.Seek(subTableOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeSubTableOffset);

                // Återställ positionen och serialisera SubTable
                writer.BaseStream.Seek(currentOffset, SeekOrigin.Begin);

                // Använd Serializer-mönstret för Type 1 och Type 4, annars kasta NotImplementedException
                switch (this.LookupType)
                {
                    case 1: // Single Substitution
                            // Since SingleSubstSubTable has two formats, we must handle serialization externally or via abstract method.
                            // Assuming we use an external Serializer for LookupType 1:
                        var singleSubstSubTable = (SingleSubstSubTable)subTable;
                        // We'll need a way to serialize based on its internal format (Format 1 or 2)

                        // Temporary fix: This needs an actual serializer class
                        // SingleSubstSubTableSerializer.Serialize(singleSubstSubTable, writer);

                        // For now, let's fall back to internal Serialize which must be implemented
                        subTable.Serialize(writer);
                        break;

                    case 4: // Ligature Substitution
                            // Vi vet att LigatureSubstSubTable implementerar Serialize (Steg 4 från tidigare svar)
                        subTable.Serialize(writer);
                        break;

                    default:
                        // Handle unsupported types if necessary
                        subTable.Serialize(writer); // Will likely throw NotImplementedException if not handled
                        break;
                }

                subIndex++;
            }
        }
    }
}
