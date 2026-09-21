/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2026         EPPlus Software AB           OpenType font implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Name;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Core scanning logic – shared between FontScannerV2 and cache.
    /// </summary>
    internal static class FontScannerV2Core
    {
        /// <summary>
        /// Scans a single TTF/OTF face at the given offset.
        /// Called only by FontScannerCache.
        /// </summary>
        internal static FontFaceInfo ScanSingleFace(string filePath, long offset)
        {
            var info = new FontFaceInfo
            {
                FilePath = filePath,
                OffsetInFile = offset,
                LastWriteTimeUtc = File.GetLastWriteTimeUtc(filePath),
                TableRecords = new Dictionary<string, TableRecord>(StringComparer.Ordinal),
                Format = FontFormat.Ttf // default fallback
            };

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (FontsBinaryReader reader = new FontsBinaryReader(fs))
            {
                fs.Position = offset;

                uint tag = reader.ReadUInt32BigEndian();

                // 1. TTC header?
                if (tag == 0x74746366) // "ttcf"
                {
                    // This should normally be handled by ScanTtcFile, but we are defensive
                    info.Format = FontFormat.Ttc;

                    // We cannot parse table directory from here – return early with minimal info
                    // This face will be properly rescanned via ScanTtcFile if needed
                    return info;
                }

                // 2. Regular OpenType/TrueType font
                if (tag == 0x00010000) // TrueType (standard sfnt version)
                {
                    info.Format = FontFormat.Ttf;
                }
                else if (tag == 0x4F54544F) // "OTTO" – OpenType with CFF outlines
                {
                    info.Format = FontFormat.Otf;
                }
                else
                {
                    // Unknown or corrupt font – throw with clear message
                    throw new InvalidOperationException(
                        string.Format("Invalid or unsupported font file signature: 0x{0:X8} in file '{1}' at offset {2}",
                            tag, filePath, offset));
                }

                // Continue with normal parsing...
                ushort numTables = reader.ReadUInt16BigEndian();
                reader.ReadUInt16BigEndian(); // searchRange
                reader.ReadUInt16BigEndian(); // entrySelector
                reader.ReadUInt16BigEndian();     // rangeShift

                for (int i = 0; i < numTables; i++)
                {
                    long tagPos = fs.Position;
                    var record = new TableRecord
                    {
                        Tag = new Tag(reader),
                        Checksum = reader.ReadUInt32BigEndian(),
                        Offset = reader.ReadUInt32BigEndian(),
                        Length = reader.ReadUInt32BigEndian()
                    };
                    info.TableRecords[record.Tag.Value] = record;
                }

                // A font is "variable" if it carries a font variations table. We only need to know that the
                // table exists — not parse it — to decide that this face cannot be trusted to deliver a
                // non-default subfamily. No extra I/O: the table directory is already in memory.
                info.IsVariable = info.TableRecords.ContainsKey("fvar");

                if (info.TableRecords.TryGetValue("OS/2", out TableRecord os2Rec))
                {
                    // fsSelection sits at byte offset 62 in the OS/2 table, after sFamilyClass (30),
                    // panose[10] (32-41), ulUnicodeRange1-4 (42-57) and achVendID (58-61).
                    // Reading at offset 32 returns the first two PANOSE bytes instead.
                    const int fsSelectionOffset = 62;

                    // Every OS/2 version (0 and up) is at least 78 bytes, so a table too short to hold
                    // fsSelection is malformed. Check up front rather than relying on the read throwing.
                    if (os2Rec.Length >= fsSelectionOffset + 2)
                    {
                        try
                        {
                            fs.Position = info.OffsetInFile + os2Rec.Offset + fsSelectionOffset;
                            info.FsSelection = reader.ReadUInt16BigEndian();
                        }
                        catch
                        {
                            // Om tabellen är korrupt eller för kort → ignorera, behåll 0
                            info.FsSelection = 0;
                        }
                    }
                    else
                    {
                        info.FsSelection = 0;
                    }
                }
                else
                {
                    info.FsSelection = 0;
                }

                // Read name table for family/subfamily
                if (info.TableRecords.TryGetValue("name", out TableRecord nameRec))
                {
                    fs.Position = nameRec.Offset;
                    byte[] nameBytes = reader.ReadBytes((int)nameRec.Length);

                    var nameTable = new NameTable();
                    nameTable.Os2FsSelection = info.FsSelection;
                    nameTable.LoadFromBytes(nameBytes);

                    info.FamilyName = nameTable.GetFamilyName() ?? "Unknown Family";
                    info.SubfamilyName = nameTable.GetSubfamilyName() ?? "Regular";
                    info.Subfamily = nameTable.GetSubfamilyEnum();
                    info.FullFontName = nameTable.GetFullFontName() ?? "Unknown Family Regular";
                }
                else
                {
                    info.FamilyName = "Unknown Family";
                    info.SubfamilyName = "Regular";
                    info.Subfamily = FontSubFamily.Regular;
                }
            }

            return info;
        }
    }
}
