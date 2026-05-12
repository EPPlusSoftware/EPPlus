/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/
  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/02/2026         EPPlus Software AB           TTC support (extracted from DefaultFontResolver)
  05/06/2026         EPPlus Software AB           Moved into IFontFileReader implementation
 *************************************************************************************************/
using System.IO;

namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Default file-system implementation of IFontFileReader.
    /// Reads font bytes directly from disk. TTC (TrueType Collection) files are handled
    /// transparently — the correct font is extracted and returned as standalone TTF bytes.
    /// </summary>
    internal sealed class DefaultFontFileReader : IFontFileReader
    {
        public byte[] ReadFontBytes(FontFaceInfo face)
        {
            byte[] fileBytes = File.ReadAllBytes(face.FilePath);

            if (face.OffsetInFile > 0)
                return ExtractTtcFont(fileBytes, face.OffsetInFile);

            return fileBytes;
        }

        /// <summary>
        /// Extracts a single TTF font from a TTC collection file.
        /// Reads the font's table directory at the given offset, copies all referenced table data,
        /// and returns a new valid standalone TTF byte array with recalculated offsets.
        /// </summary>
        /// <param name="ttcBytes">Raw bytes of the TTC file.</param>
        /// <param name="offset">Byte offset within the TTC file where the target font's SFNT header begins.</param>
        /// <returns>Standalone TTF bytes for the font at the given offset.</returns>
        private static byte[] ExtractTtcFont(byte[] ttcBytes, long offset)
        {
            using (var ms = new MemoryStream(ttcBytes))
            using (var reader = new BinaryReader(ms))
            {
                // Read SFNT header at offset
                ms.Position = offset;
                uint sfntVersion = ReadUInt32BE(reader);
                ushort numTables = ReadUInt16BE(reader);
                ms.Position += 6; // skip searchRange, entrySelector, rangeShift

                // Read table records from the font's table directory
                var tags = new string[numTables];
                var checksums = new uint[numTables];
                var srcOffsets = new uint[numTables];
                var lengths = new uint[numTables];

                for (int i = 0; i < numTables; i++)
                {
                    byte[] tagBytes = reader.ReadBytes(4);
                    tags[i] = new string(new char[] { (char)tagBytes[0], (char)tagBytes[1], (char)tagBytes[2], (char)tagBytes[3] });
                    checksums[i] = ReadUInt32BE(reader);
                    srcOffsets[i] = ReadUInt32BE(reader);
                    lengths[i] = ReadUInt32BE(reader);
                }

                // Calculate output size: SFNT header (12) + table directory (16 * n) + aligned table data
                int directorySize = 12 + numTables * 16;
                int totalSize = directorySize;
                for (int i = 0; i < numTables; i++)
                    totalSize += (int)((lengths[i] + 3u) & ~3u);

                byte[] result = new byte[totalSize];

                using (var outMs = new MemoryStream(result))
                using (var writer = new BinaryWriter(outMs))
                {
                    // Write SFNT header
                    WriteUInt32BE(writer, sfntVersion);
                    WriteUInt16BE(writer, numTables);

                    // Recompute searchRange, entrySelector, rangeShift per spec
                    int pot = 1, log2 = 0;
                    while (pot * 2 <= numTables) { pot *= 2; log2++; }
                    WriteUInt16BE(writer, (ushort)(pot * 16));   // searchRange
                    WriteUInt16BE(writer, (ushort)log2);         // entrySelector
                    WriteUInt16BE(writer, (ushort)((numTables - pot) * 16)); // rangeShift

                    // Calculate new table offsets (table data starts right after directory)
                    uint[] newOffsets = new uint[numTables];
                    uint newOffset = (uint)directorySize;
                    for (int i = 0; i < numTables; i++)
                    {
                        newOffsets[i] = newOffset;
                        newOffset += (lengths[i] + 3u) & ~3u;
                    }

                    // Write table directory with updated offsets
                    for (int i = 0; i < numTables; i++)
                    {
                        foreach (char c in tags[i]) writer.Write((byte)c);
                        WriteUInt32BE(writer, checksums[i]);
                        WriteUInt32BE(writer, newOffsets[i]);
                        WriteUInt32BE(writer, lengths[i]);
                    }

                    // Copy table data from TTC into the new buffer at recalculated offsets
                    for (int i = 0; i < numTables; i++)
                    {
                        System.Array.Copy(ttcBytes, srcOffsets[i], result, newOffsets[i], lengths[i]);
                    }
                }

                return result;
            }
        }

        private static uint ReadUInt32BE(BinaryReader r)
        {
            byte[] b = r.ReadBytes(4);
            return (uint)(b[0] << 24 | b[1] << 16 | b[2] << 8 | b[3]);
        }

        private static ushort ReadUInt16BE(BinaryReader r)
        {
            byte[] b = r.ReadBytes(2);
            return (ushort)(b[0] << 8 | b[1]);
        }

        private static void WriteUInt32BE(BinaryWriter w, uint v)
        {
            w.Write((byte)(v >> 24));
            w.Write((byte)(v >> 16));
            w.Write((byte)(v >> 8));
            w.Write((byte)v);
        }

        private static void WriteUInt16BE(BinaryWriter w, ushort v)
        {
            w.Write((byte)(v >> 8));
            w.Write((byte)v);
        }
    }
}