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
using EPPlus.Fonts.OpenType.FontLocalization;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public ushort Os2FsSelection { get; internal set; }


        /// <summary>
        /// Deep clone of NameTable. 
        /// Header fields (format, count, stringOffset) will be recomputed during serialization.
        /// </summary>
        public NameTable Clone()
        {
            var clone = new NameTable
            {
                // Keep the current format; will typically be 0 and recomputed anyway.
                format = this.format,

                // 'count' and 'stringOffset' are recomputed in SerializeInternal
                count = this.count,
                stringOffset = this.stringOffset,

                NameRecords = CloneRecords(this.NameRecords)
            };
            return clone;
        }

        private static NameRecord[] CloneRecords(NameRecord[] source)
        {
            if (source == null || source.Length == 0)
                return new NameRecord[0];

            var copy = new NameRecord[source.Length];
            for (int i = 0; i < source.Length; i++)
                copy[i] = source[i]?.Clone();
            return copy;
        }

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            // Step 1: Write header
            format = 0;
            count = (ushort)(NameRecords?.Length ?? 0);
            stringOffset = (ushort)(6 + count * 12); // 6 bytes header + 12 bytes per record

            writer.WriteUInt16BigEndian(format);
            writer.WriteUInt16BigEndian(count);
            writer.WriteUInt16BigEndian(stringOffset);

            // Step 2: Prepare string data with ENCODING-AWARE deduplication
            var stringData = new List<byte>();
            // FIX: Include encoding in the key to avoid collisions between different encodings
            var stringOffsetMap = new Dictionary<string, ushort>();

            foreach (var record in NameRecords)
            {
                var encoding = GetEncodingForRecord(record);
                var str = record.Name ?? string.Empty;
                var encoded = encoding.GetBytes(str);

                // FIX: Create a unique key that includes both the text AND the encoding
                // This ensures that "Aptos Narrow" in ISO-8859-1 and UTF-16BE get different offsets
                string dedupKey = $"{str}|{encoding.CodePage}";

                if (!stringOffsetMap.TryGetValue(dedupKey, out var offset))
                {
                    offset = (ushort)stringData.Count;
                    stringOffsetMap[dedupKey] = offset;
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

        // ALSO UPDATE GetEncodingForRecord() TO RESPECT encodingId:
        private static Encoding GetEncodingForRecord(NameRecord record)
        {
            // Unicode platform
            if (record.platformId == 0)
            {
                return Encoding.BigEndianUnicode; // UTF-16BE
            }

            // Macintosh platform
            if (record.platformId == 1)
            {
                // encodingId 0 = MacRoman (single-byte)
                // .NET doesn't have MacRoman (codepage 10000) in all versions
                // Fallback to ISO-8859-1 (Latin-1)
                try
                {
                    return Encoding.GetEncoding("ISO-8859-1");
                }
                catch
                {
                    return Encoding.UTF8;
                }
            }

            // Windows platform - CRITICAL: Must respect encodingId!
            if (record.platformId == 3)
            {
                return GetWindowsEncoding(record.encodingId);
            }

            // Fallback for unknown platforms
            return Encoding.UTF8;
        }
        private static readonly Dictionary<ushort, Encoding> _fallbackEncodings = new Dictionary<ushort, Encoding>();
        // ADD GetWindowsEncoding() method:
        private static Encoding GetWindowsEncoding(ushort encodingId)
        {
            try
            {
                if(_fallbackEncodings.TryGetValue(encodingId, out var cached))
                {
                    return cached;
                }
                switch (encodingId)
                {
                    case 0:
                        // Symbol encoding – fallback to Windows-1252
                        return Encoding.GetEncoding(1252);

                    case 1:
                        // Unicode BMP (UCS-2) – UTF-16BE
                        // THIS IS THE MOST COMMON for modern fonts!
                        return Encoding.BigEndianUnicode;

                    case 2:
                        // Shift-JIS – Japanese
                        return Encoding.GetEncoding(932);

                    case 3:
                        // GB2312 – Simplified Chinese
                        return Encoding.GetEncoding(936);

                    case 4:
                        // Big5 – Traditional Chinese
                        return Encoding.GetEncoding(950);

                    case 5:
                        // Wansung – Korean
                        return Encoding.GetEncoding(949);

                    case 6:
                        // Johab – Korean (alternative)
                        return Encoding.GetEncoding(1361);

                    case 10:
                        // Unicode full repertoire
                        // .NET doesn't have UTF-32BE directly, use UTF-16BE
                        return Encoding.BigEndianUnicode;

                    default:
                        // Fallback for unknown encodings
                        return Encoding.BigEndianUnicode;
                }
            }
            catch (NotSupportedException)
            {
                // If encoding doesn't exist (e.g. in .NET Standard without System.Text.Encoding.CodePages)
                // Fallback to UTF-16BE - this is safe for most modern fonts
                lock (_fallbackEncodings)
                {
                    _fallbackEncodings.Add(encodingId, Encoding.BigEndianUnicode);
                }

                return Encoding.BigEndianUnicode;
            }
            catch (ArgumentException)
            {
                return Encoding.BigEndianUnicode;
            }
        }

        //private Encoding GetEncodingForRecord(NameRecord record)
        //{
        //    if (record.platformId == 0)
        //        return Encoding.GetEncoding("utf-16BE");

        //    if (record.platformId == 1)
        //        return Encoding.GetEncoding(10000); // MacRoman

        //    if (record.platformId == 3)
        //        return NameTableLoader.GetWindowsEncoding(record.encodingId);

        //    return Encoding.UTF8; // Fallback
        //}

        /// <summary>
        /// Loads the name table from raw table bytes (used during font scanning).
        /// Only populates NameRecords – header fields are ignored as they will be rebuilt on write.
        /// </summary>
        /// <param name="tableBytes">Complete content of the 'name' table</param>
        /// <summary>
        /// Loads the name table from raw table bytes (used during font scanning).
        /// Populates NameRecords with decoded strings. Header fields are preserved for compatibility.
        /// </summary>
        /// <param name="tableBytes">Complete content of the 'name' table</param>
        /// <summary>
        /// Loads the name table from raw table bytes (used during font scanning).
        /// Populates NameRecords with decoded strings. Header fields are preserved for compatibility.
        /// </summary>
        /// <param name="tableBytes">Complete content of the 'name' table</param>
        public void LoadFromBytes(byte[] tableBytes)
        {
            if (tableBytes == null || tableBytes.Length < 6)
                throw new ArgumentException("Invalid name table data");

            using (var ms = new MemoryStream(tableBytes))
            using (var reader = new FontsBinaryReader(ms))
            {
                format = reader.ReadUInt16BigEndian();
                count = reader.ReadUInt16BigEndian();
                stringOffset = reader.ReadUInt16BigEndian();

                NameRecords = new NameRecord[count];

                // Step 1: Read all name record headers
                for (int i = 0; i < count; i++)
                {
                    var record = new NameRecord
                    {
                        platformId = reader.ReadUInt16BigEndian(),
                        encodingId = reader.ReadUInt16BigEndian(),
                        languageID = reader.ReadUInt16BigEndian(),
                        nameId = reader.ReadUInt16BigEndian(),
                        length = reader.ReadUInt16BigEndian(),
                        offset = reader.ReadUInt16BigEndian()
                    };

                    record.RecordType = (NameRecordTypes)record.nameId;

                    NameRecords[i] = record;
                }

                // Step 2: Read all strings from string pool
                for (int i = 0; i < count; i++)
                {
                    var record = NameRecords[i];

                    if (record.length == 0)
                    {
                        record.Name = string.Empty;
                        continue;
                    }

                    long stringPos = stringOffset + record.offset;
                    if (stringPos + record.length > tableBytes.Length)
                        continue; // corrupted

                    ms.Position = stringPos;
                    byte[] stringBytes = reader.ReadBytes(record.length);

                    Encoding encoding = GetEncodingForRecord(record);
                    record.Name = encoding.GetString(stringBytes);

                    // Vi skippar LanguageMapping helt under scanning
                    // Det är bara tungt och behövs inte för att få Family/Subfamily
                    record.LanguageMapping = null;
                }
            }
        }

        /// <summary>
        /// Returns the preferred Full Font Name (name ID 4) with proper language fallback.
        /// First tries platform-specific record without language (common default),
        /// then falls back to explicit English (US).
        /// </summary>
        public string GetFullFontName()
        {
            // 1. Platform-specific record with no language tag (most common default)
            string name = GetFirstNonEmpty(NameRecordTypes.FullFontName);
            if (!string.IsNullOrEmpty(name))
                return name;

            // 2. Explicit English (US) – fallback
            return GetEnglishName(NameRecordTypes.FullFontName);
        }

        /// <summary>
        /// Returns the font family name in the legacy/RIBBI naming system (nameID 1), which pairs
        /// with GetSubfamilyName()'s nameID 2. Together they form the family+style view that Windows,
        /// GDI and Excel present, and the one FontSubFamily's four values are defined against.
        ///
        /// nameID 16 (Typographic Family) is deliberately NOT preferred, even though it is the newer
        /// field: it belongs to the *other* naming system, paired with nameID 17. Arial Black reads
        /// "Arial Black" + "Regular" as (1+2) but "Arial" + "Black" as (16+17). Preferring 16 here
        /// while GetSubfamilyName() reads 2 (or vice versa) mixes the two systems and produces a pair
        /// that exists in neither, breaking font matching. nameID 16 is used only as a last resort,
        /// for fonts that omit nameID 1 entirely.
        /// </summary>
        public string GetFamilyName()
        {
            // Legacy Family Name (1), English first: NameRecords can hold one localized nameID 1
            // per language, and GetFirstNonEmpty returns whichever comes first in file order.
            string name = GetEnglishName(NameRecordTypes.FontFamilyName);
            if (!string.IsNullOrEmpty(name))
                return name;

            name = GetFirstNonEmpty(NameRecordTypes.FontFamilyName);
            if (!string.IsNullOrEmpty(name))
                return name;

            // Last resort only: the typographic family, for fonts that omit nameID 1.
            name = GetEnglishName(NameRecordTypes.TypographicFamilyName);
            if (!string.IsNullOrEmpty(name))
                return name;

            return GetFirstNonEmpty(NameRecordTypes.TypographicFamilyName) ?? "Unknown Family";
        }

        /// <summary>
        /// Returns the subfamily name in the legacy/RIBBI naming system (nameID 2), which pairs
        /// with GetFamilyName()'s nameID 1. nameID 2 is guaranteed by the OpenType spec to be one
        /// of "Regular"/"Bold"/"Italic"/"Bold Italic", which is exactly the FontSubFamily model.
        ///
        /// nameID 17 (Typographic Subfamily) is deliberately NOT preferred: it belongs to the
        /// *other* naming system, paired with nameID 16. Arial Black reads "Arial Black" + "Regular"
        /// as (1+2) but "Arial" + "Black" as (16+17). Mixing them yields the pair "Arial Black" +
        /// "Black", which exists in neither system and matches no style request.
        ///
        /// Returns null when the font carries no subfamily name at all. Callers that need a display
        /// string apply their own default; GetSubfamilyEnum relies on null to know it should fall
        /// back to OS/2 fsSelection instead.
        /// </summary>
        public string GetSubfamilyName()
        {
            // Legacy Subfamily (2), English first: a font can carry one localized nameID 2 per
            // language, and GetFirstNonEmpty returns whichever happens to come first in file order.
            string name = GetEnglishName(NameRecordTypes.FontSubfamilyName);
            if (!string.IsNullOrEmpty(name))
                return name;

            name = GetFirstNonEmpty(NameRecordTypes.FontSubfamilyName);
            if (!string.IsNullOrEmpty(name))
                return name;

            // Last resort only: the typographic subfamily, for fonts that omit nameID 2.
            name = GetEnglishName(NameRecordTypes.TypographicSubfamilyName);
            if (!string.IsNullOrEmpty(name))
                return name;

            return GetFirstNonEmpty(NameRecordTypes.TypographicSubfamilyName);
        }

        public FontSubFamily GetSubfamilyEnum()
        {
            string subfamily = GetSubfamilyName();

            // No subfamily name in the name table at all → fall back to OS/2 fsSelection.
            if (string.IsNullOrEmpty(subfamily))
                return GetSubfamilyFromFsSelection();

            string lower = subfamily.ToLowerInvariant();

            // Exakta matchningar först
            if (lower == "regular" || lower == "normal" || lower == "roman" || lower == "book")
                return FontSubFamily.Regular;

            if (lower == "bold")
                return FontSubFamily.Bold;

            if (lower == "italic" || lower == "oblique")
                return FontSubFamily.Italic;

            if (lower.Contains("bold") && lower.Contains("italic"))
                return FontSubFamily.BoldItalic;
            if (lower.Contains("bold"))
                return FontSubFamily.Bold;
            if (lower.Contains("italic") || lower.Contains("oblique"))
                return FontSubFamily.Italic;

            // Weight names beyond "Bold" (Black, Heavy, Demi, Light, Medium, etc.) don't fit the
            // 4-value RIBBI model and are NOT treated as Bold here. Fonts using these names
            // (e.g. "Arial Black", "Segoe UI Black") already distinguish themselves via FamilyName,
            // so their base instance is Regular within FontSubFamily. Mapping them to Bold would
            // falsely disqualify an exact match against a Regular request, sending the resolver
            // into the fallback chain even though the font is installed.
            //
            // Note that we return Regular here rather than consulting fsSelection: the name table
            // did give us an answer, it just isn't expressible in four values. fsSelection is not a
            // tie-breaker for that case — vendors commonly set its BOLD bit on Black/Heavy faces as
            // a legacy hint for apps that can't read the name table, so consulting it here would
            // silently re-introduce this exact bug through a different path.
            return FontSubFamily.Regular;
        }

        private FontSubFamily GetSubfamilyFromFsSelection()
        {
            // Dessa bitar är definierade i OpenType-specen
            const ushort ITALIC = 0x0001;
            const ushort BOLD = 0x0020;
            // OBS: Bit 6 (0x0040) är UNDERSCORED, bit 9 (0x0200) är REGULAR (används ibland)

            bool isItalic = (Os2FsSelection & ITALIC) != 0;
            bool isBold = (Os2FsSelection & BOLD) != 0;

            if (isBold && isItalic) return FontSubFamily.BoldItalic;
            if (isBold) return FontSubFamily.Bold;
            if (isItalic) return FontSubFamily.Italic;
            return FontSubFamily.Regular;
        }

        private string GetFirstNonEmpty(NameRecordTypes type)
        {
            foreach (var r in NameRecords)
            {
                if (r.RecordType == type && !string.IsNullOrEmpty(r.Name))
                {
                    return r.Name;
                }
            }
            return null;
        }

        private string GetEnglishName(NameRecordTypes type)
        {
            foreach (var record in NameRecords)
            {
                if (record.RecordType == type &&
                    record.LanguageMapping != null &&
                    record.LanguageMapping.Language == Languages.English &&
                    !string.IsNullOrEmpty(record.Name))
                {
                    return record.Name;
                }
            }
            return null;
        }

        /// <summary>
        /// Returns the PostScript Name (nameID 6).
        /// Follows OpenType recommendations:
        /// 1. Prefer platform 3 (Windows) → UTF-16BE
        /// 2. Then platform 1 (Macintosh) → MacRoman
        /// 3. Then platform 0 (Unicode)
        /// If no PostScript name exists, fallback to a sanitized FullFontName.
        /// </summary>
        public string PostScriptName
        {
            get
            {
                // 1. Windows (platform 3) – most reliable
                var win = NameRecords
                    .Where(r => r.RecordType == NameRecordTypes.PostScriptName && r.platformId == 3)
                    .Select(r => r.Name)
                    .FirstOrDefault(n => !string.IsNullOrEmpty(n));
                if (!string.IsNullOrEmpty(win))
                    return SanitizePsName(win);

                // 2. Unicode (platform 0)
                var uni = NameRecords
                    .Where(r => r.RecordType == NameRecordTypes.PostScriptName && r.platformId == 0)
                    .Select(r => r.Name)
                    .FirstOrDefault(n => !string.IsNullOrEmpty(n));
                if (!string.IsNullOrEmpty(uni))
                    return SanitizePsName(uni);

                // 3. Macintosh (platform 1)
                var mac = NameRecords
                    .Where(r => r.RecordType == NameRecordTypes.PostScriptName && r.platformId == 1)
                    .Select(r => r.Name)
                    .FirstOrDefault(n => !string.IsNullOrEmpty(n));
                if (!string.IsNullOrEmpty(mac))
                    return SanitizePsName(mac);

                // 4. If nameID 6 is missing – fallback to FullFontName (Windows)
                var full = GetFullFontName();
                if (!string.IsNullOrEmpty(full))
                    return SanitizePsName(full);

                // 5. Last fallback
                return "UnknownPSName";
            }
        }
        /// <summary>
        /// Sanitizes a name so it always becomes a valid PostScript-compatible font name.
        /// Removes illegal characters and replaces whitespace with hyphens.
        /// </summary>
        private static string SanitizePsName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "UnknownPSName";

            var sb = new StringBuilder(name.Length);

            foreach (char c in name)
            {
                if (char.IsWhiteSpace(c))
                {
                    sb.Append('-');
                    continue;
                }

                // Valid ASCII range for PostScript names
                if (c >= 33 && c <= 126)
                {
                    sb.Append(c);
                    continue;
                }

                // Skip invalid characters
            }

            // If everything got stripped
            return sb.Length > 0 ? sb.ToString() : "UnknownPSName";
        }
    }
}
