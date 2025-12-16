using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LookupListTable : FontTableElement
    {
        public List<LookupTable> Lookups { get; set; } = new List<LookupTable>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long listStartOffset = writer.BaseStream.Position;

            // 1. USHORT LookupCount
            writer.WriteUInt16BigEndian((ushort)this.Lookups.Count);

            // 2. USHORT[] LookupOffsets (Placeholders)
            List<long> offsetPositions = new List<long>();
            for (int i = 0; i < this.Lookups.Count; i++)
            {
                offsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // 3. Serialize each LookupTable and update offsets
            for (int i = 0; i < this.Lookups.Count; i++)
            {
                long currentPos = writer.BaseStream.Position;
                ushort relativeOffset = (ushort)(currentPos - listStartOffset);

                // Gå tillbaka och fyll i offset
                writer.BaseStream.Seek(offsetPositions[i], SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeOffset);

                // Tillbaka till skrivposition och kör Lookupens egen Serialize
                writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);
                this.Lookups[i].Serialize(writer);
            }
        }
    }
}
