using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class ScriptTable : FontTableElement
    {
        // Property is kept mainly for deserialization purposes, its value is calculated during serialization.
        public ushort DefaultLangSysOffset { get; set; }

        /// <summary>
        /// Contains the deserialized LangSysTable object for the default language system. 
        /// Checked for null status during serialization.
        /// </summary>
        public LangSysTable DefaultLangSys { get; set; }

        /// <summary>
        /// Contains records for all language systems supported by this script (excluding the default one).
        /// </summary>
        public List<LangSysRecord> LangSysRecords { get; set; } = new List<LangSysRecord>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // The writer's current position is the start of the ScriptTable.
            long scriptTableStartOffset = writer.BaseStream.Position;

            // 1. USHORT DefaultLangSysOffset (Placeholder)
            long defaultLangSysOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 2. USHORT LangSysCount
            writer.WriteUInt16BigEndian((ushort)this.LangSysRecords.Count);

            // 3. LangSysRecord[] - Write tags and placeholders for LangSysOffset
            List<long> langSysOffsetPositions = new List<long>();

            foreach (var record in this.LangSysRecords)
            {
                // Write LangSysTag (4 bytes)
                writer.WriteUInt32BigEndian(record.LangSysTag);

                // Placeholder for LangSysOffset (2 bytes)
                langSysOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- Serialize LangSysTable(s) ---

            // 1. Serialize DefaultLangSysTable
            if (this.DefaultLangSys != null)
            {
                long currentOffset = writer.BaseStream.Position;

                // Calculate DefaultLangSys offset: LangSys start - ScriptTable start
                ushort relativeOffset = (ushort)(currentOffset - scriptTableStartOffset);

                // Save current position
                long positionBeforeSeek = writer.BaseStream.Position;

                // Fill in placeholder
                writer.BaseStream.Seek(defaultLangSysOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeOffset);

                // Restore position
                writer.BaseStream.Seek(positionBeforeSeek, SeekOrigin.Begin);

                // Serialize DefaultLangSysTable
                this.DefaultLangSys.Serialize(writer);
            }

            // 2. Serialize other LangSysTables from LangSysRecords
            int recordIndex = 0;
            foreach (var record in this.LangSysRecords)
            {
                long currentOffset = writer.BaseStream.Position;

                // Calculate offset: LangSysTable start - ScriptTable start
                ushort relativeOffset = (ushort)(currentOffset - scriptTableStartOffset);

                // Fill in placeholder for LangSysRecord
                long recordOffsetPos = langSysOffsetPositions[recordIndex];

                // Save current position
                long positionBeforeSeek = writer.BaseStream.Position;

                writer.BaseStream.Seek(recordOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeOffset);
                writer.BaseStream.Seek(positionBeforeSeek, SeekOrigin.Begin); // Restore position

                // Serialize LangSysTable
                record.LangSysTable.Serialize(writer);
                recordIndex++;
            }
        }
    }
}
