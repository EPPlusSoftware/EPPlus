using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class ScriptListTable : FontTableElement
    {
        public List<ScriptRecord> ScriptRecords { get; set; } = new List<ScriptRecord>();

        // Implementerar den korrekta abstrakta metoden
        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. USHORT ScriptCount
            writer.WriteUInt16BigEndian((ushort)this.ScriptRecords.Count);

            // Placeholder for USHORT[] ScriptRecords (ScriptTag + ScriptTableOffset)
            List<long> recordOffsetPositions = new List<long>();
            long scriptListStartOffset = writer.BaseStream.Position - sizeof(ushort);

            foreach (var record in this.ScriptRecords)
            {
                // Write ScriptTag (4 bytes)
                writer.Write(record.ScriptTag.ToBytes());

                // Placeholder for ScriptTableOffset (2 bytes)
                recordOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- Skriv ut ScriptTables ---

            // Vi behöver en sorterad lista över ScriptRecords för att hantera offsets korrekt.
            // Då ScriptRecords redan är en lista, itererar vi över den (och antar att den är i rätt ordning efter CreateSubset).

            int recordIndex = 0;
            foreach (var record in this.ScriptRecords)
            {
                long currentOffset = writer.BaseStream.Position;

                // Beräkna offset: ScriptTable start - ScriptList start
                ushort relativeScriptTableOffset = (ushort)(currentOffset - scriptListStartOffset);

                // 1. Gå tillbaka och fyll i offseten i ScriptRecord
                long recordOffsetPos = recordOffsetPositions[recordIndex];
                writer.BaseStream.Seek(recordOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeScriptTableOffset);

                // 2. Återställ positionen och serialisera ScriptTable
                writer.BaseStream.Seek(currentOffset, SeekOrigin.Begin);

                // Serialisera ScriptTable
                // OBS: Vi antar att varje ScriptRecord har en referens till sin ScriptTable.
                // record.ScriptTable must implement Serialize
                record.ScriptTable.Serialize(writer);

                recordIndex++;
            }
        }
    }
}
