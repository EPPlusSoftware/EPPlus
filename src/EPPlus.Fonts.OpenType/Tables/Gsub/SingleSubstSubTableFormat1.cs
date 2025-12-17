using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
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
            int index = Coverage.GetGlyphIndex(baseGlyphId);
            if (index == -1) return 0; // Not covered

            // Format 1 adds a delta to the original GID
            // (ushort wrap-around is handled automatically by C# casting)
            return (ushort)(baseGlyphId + DeltaGlyphID);
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Spara starten av denna subtable direkt
            long subTableStart = writer.BaseStream.Position;

            // 1. USHORT SubtableFormat (1)
            writer.WriteUInt16BigEndian(1);

            // 2. USHORT CoverageOffset (Placeholder)
            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 3. SSHORT DeltaGlyphID
            writer.WriteInt16BigEndian(this.DeltaGlyphID);

            // --- Skriv ut CoverageTable ---
            if (this.Coverage != null)
            {
                // Använd verktygsmetoden för att skriva offseten automatiskt
                this.WriteRelativeOffset(writer, subTableStart, covOffsetPos);

                // Låt CoverageTable sköta sin egen serialisering (som vi satte upp tidigare)
                this.Coverage.Serialize(writer);
            }
        }
    }
}
