using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
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
            int index = Coverage.GetGlyphIndex(baseGlyphId);
            if (index == -1 || index >= SubstituteGlyphIDs.Length) return 0;

            // Format 2 maps the coverage index to the substitute array
            return SubstituteGlyphIDs[index];
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long startPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(2); // Format 2

            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Placeholder

            writer.WriteUInt16BigEndian((ushort)this.SubstituteGlyphIDs.Length);
            foreach (var gid in this.SubstituteGlyphIDs)
                writer.WriteUInt16BigEndian(gid);

            // Skriv Coverage och uppdatera offset
            this.WriteRelativeOffset(writer, startPos, covOffsetPos);
            this.Coverage.Serialize(writer);
        }
    }
}
