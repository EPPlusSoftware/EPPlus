/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Name
{
    public class NameTable : FontTableBase
    {
        public override string Name => TableNames.Name;

        public override bool IsEssentialTable => false;
        public ushort format { get; set; }

        public ushort count { get; set; }

        public ushort stringOffset { get; set; }

        public NameRecord[] NameRecords { get; set; }

        internal override void Clear()
        {
            throw new System.NotImplementedException();
        }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            // Step 1: Write header
            format = 0;
            count = (ushort)(NameRecords?.Length ?? 0);
            stringOffset = (ushort)(6 + count * 12); // 6 bytes header + 12 bytes per record

            writer.WriteUInt16BigEndian(format);
            writer.WriteUInt16BigEndian(count);
            writer.WriteUInt16BigEndian(stringOffset);


            // Step 2: Prepare string data with deduplication
            var stringData = new List<byte>();
            var stringOffsetMap = new Dictionary<string, ushort>();

            foreach (var record in NameRecords)
            {
                var encoding = GetEncodingForRecord(record);
                var str = record.Name ?? string.Empty;
                var encoded = encoding.GetBytes(str);

                if (!stringOffsetMap.TryGetValue(str, out var offset))
                {
                    offset = (ushort)stringData.Count;
                    stringOffsetMap[str] = offset;
                    stringData.AddRange(encoded);
                }

                record.length = (ushort)encoded.Length;
                record.offset = offset;
            }


            // Step 3: Write NameRecords
            foreach (var record in NameRecords)
            {
                record.Serialize(writer);
            }

            // Step 4: Write string pool
            writer.Write(stringData.ToArray());
        }


        private Encoding GetEncodingForRecord(NameRecord record)
        {
            if (record.platformId == 0)
                return Encoding.GetEncoding("utf-16BE");

            if (record.platformId == 1)
                return Encoding.GetEncoding(10000); // MacRoman

            if (record.platformId == 3)
                return NameTableLoader.GetWindowsEncoding(record.encodingId);

            return Encoding.UTF8; // Fallback
        }
    }
}
