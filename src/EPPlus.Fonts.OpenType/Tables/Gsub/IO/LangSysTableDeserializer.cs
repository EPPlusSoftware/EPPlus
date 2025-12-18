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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class LangSysTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public LangSysTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public LangSysTable Deserialize(long langSysTableStartOffset)
        {
            // Set reader position to the start of the LangSysTable
            _reader.BaseStream.Seek(langSysTableStartOffset, SeekOrigin.Begin);

            LangSysTable table = new LangSysTable();

            // 1. USHORT LookupOrder (Reserved, set to 0)
            table.LookupOrder = _reader.ReadUInt16BigEndian();

            // 2. USHORT RequiredFeatureIndex (Index into FeatureList. 0xFFFF if none required)
            table.RequiredFeatureIndex = _reader.ReadUInt16BigEndian();

            // 3. USHORT FeatureIndexCount
            ushort count = _reader.ReadUInt16BigEndian();
            table.FeatureIndexCount = count; // Property only exists for deserialization tracking/debugging

            // 4. USHORT[] FeatureIndices
            table.FeatureIndices = new ushort[count];
            for (int i = 0; i < count; i++)
            {
                // Read USHORT FeatureIndex
                table.FeatureIndices[i] = _reader.ReadUInt16BigEndian();
            }

            return table;
        }
    }
}
