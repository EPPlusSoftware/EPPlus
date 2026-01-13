/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           Shared ScriptList loader
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts
{
    /// <summary>
    /// Shared loader for ScriptListTable used by both GSUB and GPOS
    /// </summary>
    internal static class ScriptListTableLoader
    {
        public static ScriptListTable Load(FontsBinaryReader reader, long scriptListStart)
        {
            var scriptList = new ScriptListTable();

            ushort scriptCount = reader.ReadUInt16BigEndian();

            // Read script records
            var scriptOffsets = new List<ScriptOffsetRecord>();
            for (int i = 0; i < scriptCount; i++)
            {
                var tag = new Tag(reader);
                ushort offset = reader.ReadUInt16BigEndian();
                scriptOffsets.Add(new ScriptOffsetRecord { Tag = tag, Offset = offset });
            }

            long positionAfterRecords = reader.BaseStream.Position;

            // Load script tables
            foreach (var record in scriptOffsets)
            {
                // ✅ Seek to script table position
                reader.BaseStream.Seek(scriptListStart + record.Offset, SeekOrigin.Begin);

                // ✅ LoadScriptTable reads from current position
                var scriptTable = LoadScriptTable(reader);

                scriptList.ScriptRecords.Add(new ScriptRecord
                {
                    ScriptTag = record.Tag,
                    ScriptOffset = record.Offset,
                    ScriptTable = scriptTable
                });
            }

            reader.BaseStream.Seek(positionAfterRecords, SeekOrigin.Begin);
            return scriptList;
        }

        private static ScriptTable LoadScriptTable(FontsBinaryReader reader)
        {
            long scriptTableStart = reader.BaseStream.Position;

            var scriptTable = new ScriptTable();
            ushort defaultLangSysOffset = reader.ReadUInt16BigEndian();
            scriptTable.DefaultLangSysOffset = defaultLangSysOffset;
            ushort langSysCount = reader.ReadUInt16BigEndian();

            var recordsToLoad = new Dictionary<uint, ushort>();
            for (int i = 0; i < langSysCount; i++)
            {
                uint langSysTag = reader.ReadUInt32BigEndian();
                ushort langSysOffset = reader.ReadUInt16BigEndian();

                if (!recordsToLoad.ContainsKey(langSysTag))
                {
                    recordsToLoad.Add(langSysTag, langSysOffset);
                }
            }

            long positionAfterRecords = reader.BaseStream.Position;
            var langSysDeserializer = new LangSysTableDeserializer(reader);

            // Load default LangSys
            if (defaultLangSysOffset > 0)
            {
                long langSysAbsoluteStart = scriptTableStart + defaultLangSysOffset;
                scriptTable.DefaultLangSys = langSysDeserializer.Deserialize(langSysAbsoluteStart);
            }

            // Load other LangSys records
            foreach (var kvp in recordsToLoad)
            {
                long langSysAbsoluteStart = scriptTableStart + kvp.Value;

                var langSysTable = langSysDeserializer.Deserialize(langSysAbsoluteStart);

                scriptTable.LangSysRecords.Add(new LangSysRecord
                {
                    LangSysTag = kvp.Key,
                    LangSysTable = langSysTable
                });
            }

            reader.BaseStream.Seek(positionAfterRecords, SeekOrigin.Begin);

            return scriptTable;
        }

        // Helper struct for .NET 3.5 compatibility (no tuples)
        private struct ScriptOffsetRecord
        {
            public Tag Tag;
            public ushort Offset;
        }
    }
}