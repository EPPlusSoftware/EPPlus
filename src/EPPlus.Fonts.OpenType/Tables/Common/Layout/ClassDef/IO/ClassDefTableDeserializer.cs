/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  16/01/2026         EPPlus Software AB           ClassDef deserializer (Format 1 & 2)
 *************************************************************************************************/
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef.IO
{
    /// <summary>
    /// Deserializes ClassDef tables (Format 1 and 2) from OpenType fonts.
    /// Shared between GSUB and GPOS.
    /// </summary>
    internal class ClassDefTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public ClassDefTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        /// <summary>
        /// Deserializes a ClassDef table at the given absolute offset.
        /// </summary>
        /// <param name="classDefStart">Absolute byte offset where the ClassDef table starts.</param>
        public ClassDefTable Deserialize(long classDefStart)
        {
            _reader.BaseStream.Seek(classDefStart, SeekOrigin.Begin);

            ushort format = _reader.ReadUInt16BigEndian();

            if (format == 1)
            {
                return ReadFormat1(classDefStart);
            }
            else if (format == 2)
            {
                return ReadFormat2(classDefStart);
            }

            // Okänt format – returnera null eller kasta om du vill vara strikt
            return null;
        }

        private ClassDefTable ReadFormat1(long classDefStart)
        {
            // Format 1:
            // USHORT ClassFormat (already read)
            // USHORT StartGlyphID
            // USHORT GlyphCount
            // USHORT ClassValueArray[GlyphCount]

            var table = new ClassDefFormat1
            {
                StartGlyphID = _reader.ReadUInt16BigEndian(),
                GlyphCount = _reader.ReadUInt16BigEndian()
            };

            table.ClassValueArray = new ushort[table.GlyphCount];
            for (int i = 0; i < table.GlyphCount; i++)
            {
                table.ClassValueArray[i] = _reader.ReadUInt16BigEndian();
            }

            return table;
        }

        private ClassDefTable ReadFormat2(long classDefStart)
        {
            // Format 2:
            // USHORT ClassFormat (already read)
            // USHORT ClassRangeCount
            // ClassRangeRecord ClassRangeRecord[ClassRangeCount]

            var table = new ClassDefFormat2();

            ushort classRangeCount = _reader.ReadUInt16BigEndian();
            table.ClassRangeRecords = new List<ClassRangeRecord>(classRangeCount);

            for (int i = 0; i < classRangeCount; i++)
            {
                var rec = new ClassRangeRecord
                {
                    StartGlyphID = _reader.ReadUInt16BigEndian(),
                    EndGlyphID = _reader.ReadUInt16BigEndian(),
                    Class = _reader.ReadUInt16BigEndian()
                };

                table.ClassRangeRecords.Add(rec);
            }

            return table;
        }
    }
}
