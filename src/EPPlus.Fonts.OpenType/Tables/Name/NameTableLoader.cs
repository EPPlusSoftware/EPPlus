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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Name
{
    internal class NameTableLoader : TableLoader<NameTable>
    {
        public NameTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Name)
        {
#if NET5_0_OR_GREATER
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);  //Add Support for codepage 1252
#endif
        }

        protected override NameTable LoadInternal()
        {
            ushort format = _reader.ReadUInt16BigEndian();
            ushort count = _reader.ReadUInt16BigEndian();
            ushort stringOffset = _reader.ReadUInt16BigEndian();
            var globalStringOffset = _offset + stringOffset; 
            var records = new List<NameRecord>();
            for(var x = 0; x < count; x++)
            {
                var platformId = _reader.ReadUInt16BigEndian();
                var encodingID = _reader.ReadUInt16BigEndian();
                var languageID = _reader.ReadUInt16BigEndian();
                var nameID = _reader.ReadUInt16BigEndian();
                var length = _reader.ReadUInt16BigEndian();
                var offset = _reader.ReadUInt16BigEndian();
                var record = new NameRecord
                {
                    platformId = platformId,
                    encodingId = encodingID,
                    languageID = languageID,
                    nameId = nameID,
                    RecordType = (NameRecordTypes)nameID,
                    length = length,
                    offset = offset
                };
                records.Add(record);
            }
           
            foreach (var record in records)
            {
                _reader.BaseStream.Position = globalStringOffset + record.offset;
                var bytes = _reader.ReadBytes(record.length);
                SetName(record, bytes);
            }
            
            return new NameTable
            {
                format = format,
                count = count,
                stringOffset = stringOffset,
                NameRecords = records.ToArray()
            };
        }

        private void SetName(NameRecord record, byte[] bytes)
        {
            // Macintosh platform
            if (record.platformId == 1)
            {
                if (!Encoding.GetEncodings().Any(x => x.CodePage == 10000)) return;
                var enc = Encoding.GetEncoding(10000);
                record.Name = enc.GetString(bytes);
                if(MacintoshLanguageMappings.Mappings.ContainsKey(record.languageID))
                {
                    record.LanguageMapping = MacintoshLanguageMappings.Mappings[record.languageID];
                }
            }
            // Unicode platform
            else if (record.platformId == 0)
            {
                record.Name = Encoding.GetEncoding("utf-16BE").GetString(bytes);
            }
            // Windows platform
            else if (record.platformId == 3)
            {
                if (WindowsLanguageMappings.Mappings.ContainsKey((int)record.languageID))
                {
                    record.LanguageMapping = WindowsLanguageMappings.Mappings[record.languageID];
                }
                var enc = GetWindowsEncoding(record.encodingId);
                record.Name = enc.GetString(bytes);
            }
        }

        internal static Encoding GetWindowsEncoding(int encodingId)
        {
            // Ensure CodePagesEncodingProvider is registered before calling this method:
            // Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            try
            {
                switch (encodingId)
                {
                    case 0:
                        // Symbol encoding – not Unicode. Fallback to Windows-1252 for compatibility.
                        return Encoding.GetEncoding(1252);

                    case 1:
                        // Unicode BMP (UCS-2) – typically stored as UTF-16BE in font files.
                        return Encoding.GetEncoding("utf-16BE");

                    case 2:
                        // Shift-JIS – used for Japanese.
                        return Encoding.GetEncoding(932);

                    case 3:
                        // PRC – Simplified Chinese (GB2312).
                        return Encoding.GetEncoding(936);

                    case 4:
                        // Big5 – Traditional Chinese.
                        return Encoding.GetEncoding(950);

                    case 5:
                        // Wansung – Korean.
                        return Encoding.GetEncoding(949);

                    case 6:
                        // Johab – Korean (alternative encoding).
                        return Encoding.GetEncoding(1361);

                    case 10:
                        // Unicode full repertoire – UTF-32BE.
                        return Encoding.GetEncoding("utf-32BE");

                    default:
                        // Fallback to UTF-16LE (default in .NET) for unknown or unsupported encodings.
                        return Encoding.Unicode;
                }
            }
            catch (NotSupportedException)
            {
                // If the requested encoding is not available, fallback to UTF-16LE.
                return Encoding.Unicode;
            }
        }
    }
}
