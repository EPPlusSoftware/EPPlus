using EPPlus.Fonts.OpenType.Tables.Gsub.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class SingleSubstSubTableFormat1 : SingleSubstSubTable
    {
        public short DeltaGlyphID { get; set; } // SSHORT

        // Kräver att vi har CoverageTable (och dess index)
        public override ushort GetSubstitution(ushort baseGlyphId)
        {
            // Vi hoppar över denna komplexa lookup-logik just nu och fokuserar på läsaren.
            throw new NotImplementedException();
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. USHORT SubtableFormat (1)
            writer.WriteUInt16BigEndian(1);

            // 2. USHORT CoverageOffset (Placeholder: 2 bytes)
            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 3. SSHORT DeltaGlyphID
            writer.WriteInt16BigEndian(this.DeltaGlyphID);

            // --- Skriv ut CoverageTable ---
            long coverageStartPos = writer.BaseStream.Position;

            if (this.Coverage != null)
            {
                // Fyll i den relativa offseten. Offseten är relativ till SubTable start (dvs. 4 bytes före coverageOffsetPos)
                ushort relativeCoverageOffset = (ushort)(coverageStartPos - (coverageOffsetPos - sizeof(ushort) * 2));

                // Spara nuvarande position
                long currentPos = writer.BaseStream.Position;

                // Gå tillbaka och skriv offseten
                writer.BaseStream.Seek(coverageOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeCoverageOffset);

                // Återgå till skrivposition
                writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);

                // Serialisera Coverage
                if (this.Coverage.CoverageFormat == 1)
                {
                    new CoverageTableFormat1Serializer().Serialize((CoverageTableFormat1)this.Coverage, writer);
                }
                // Add Format 2 serialization here if needed
            }
        }
    }
}
