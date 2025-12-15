using EPPlus.Fonts.OpenType.Tables.Gsub.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class SingleSubstSubTableFormat2 : SingleSubstSubTable
    {
        public ushort GlyphCount { get; set; }
        public ushort[] SubstituteGlyphIDs { get; set; } // USHORT[]

        public override ushort GetSubstitution(ushort baseGlyphId)
        {
            throw new NotImplementedException();
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. USHORT SubtableFormat (2)
            writer.WriteUInt16BigEndian(2);

            // 2. USHORT CoverageOffset (Placeholder: 2 bytes)
            // Offseten måste beräknas relativt till SubTablens början.
            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 3. USHORT GlyphCount
            writer.WriteUInt16BigEndian(this.GlyphCount);

            // 4. USHORT[] SubstituteGlyphIDs
            foreach (ushort gid in this.SubstituteGlyphIDs)
            {
                writer.WriteUInt16BigEndian(gid);
            }

            // --- Skriv ut CoverageTable ---
            long coverageStartPos = writer.BaseStream.Position;

            if (this.Coverage != null)
            {
                // Beräkna den relativa offseten.
                // SubTable start = (coverageOffsetPos - 2 bytes för Format)
                // Längd till coverageStartPos = coverageStartPos - SubTable start.
                // Offseten är (coverageStartPos - 4 bytes för (Format + CoverageOffset))
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
