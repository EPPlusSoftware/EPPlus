using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    /// <summary>
    /// This table defines the mapping of character codes to the glyph index values used in the font. It may contain more than one subtable, in order to support more than one character encoding scheme.
    /// </summary>
    public class CmapTable : FontTableBase
    {
        internal CmapTable()
        {
            EncodingRecords = new List<EncodingRecord>();
            SubTables = new List<CmapSubtableBase>();
        }
        /// <summary>
        /// Table version number (0).
        /// </summary>
        public ushort Version { get; set; }

        /// <summary>
        /// Number of encoding tables that follow.
        /// </summary>
        public ushort NumTables { get; set; }

        /// <summary>
        /// The array of encoding records specifies particular encodings and the offset to the subtable for each encoding.
        /// </summary>
        public List<EncodingRecord> EncodingRecords { get; private set; }

        /// <summary>
        /// Array of Subtables
        /// </summary>
        public List<CmapSubtableBase> SubTables { get; private set; }


        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            // Write version and number of encoding records
            writer.WriteUInt16BigEndian(Version);
            writer.WriteUInt16BigEndian((ushort)EncodingRecords.Count);

            // Reserve space for encoding records (each is 8 bytes)
            long encodingRecordStart = writer.BaseStream.Position;
            foreach (var _ in EncodingRecords)
            {
                writer.Write(new byte[8]); // placeholder
            }

            // Serialize each subtable to a buffer and track offsets
            var subtableBuffers = new List<byte[]>();
            var subtableOffsets = new List<uint>();
            long subtableStart = writer.BaseStream.Position;

            foreach (var subtable in SubTables)
            {
                byte[] data = subtable.Serialize();
                subtableBuffers.Add(data);
                subtableOffsets.Add((uint)(writer.BaseStream.Position - encodingRecordStart));
                writer.Write(data);
            }

            // Go back and write encoding records with correct offsets
            long currentPos = writer.BaseStream.Position;
            writer.BaseStream.Seek(encodingRecordStart, SeekOrigin.Begin);

            for (int i = 0; i < EncodingRecords.Count; i++)
            {
                var record = EncodingRecords[i];
                writer.WriteUInt16BigEndian((ushort)record.PlatformId);
                writer.WriteUInt16BigEndian(record.EncodingId);
                writer.WriteUInt32BigEndian(subtableOffsets[i]);
            }

            // Return to end of stream
            writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);

        }

        public CmapSubtable4 GetSubtable4(bool throwExceptionIfNull = true)
        {
            var enc = EncodingRecords.FirstOrDefault(er => er.PlatformId == Platforms.Windows && er.EncodingId == 1);
            if (enc == null)
            {
                if(throwExceptionIfNull)
                {
                    throw new Exception("Could not find Microsoft Unicode cmap (PlatformID 3, EncodingID 1).");
                }
                return null;
            }
            var subTableIx = EncodingRecords.IndexOf(enc);
            return SubTables[subTableIx] as CmapSubtable4;
        }

    }
}
