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
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;

namespace EPPlus.Fonts.OpenType.Scanner
{
    [DebuggerDisplay("Fontfamily: {FontFamilyName} {FontSubFamilyName}")]
    public class ScannedFont : IScannedFont
    {
        internal ScannedFont(FontsBinaryReader reader, FontFormat format, string filePath)
            : this(reader, format, filePath, -1)
        {
        }

        internal ScannedFont(FontsBinaryReader reader, FontFormat format, string filePath, long offset)
        {
            _reader = reader;
            _format = format;
            _tableRecords = new Dictionary<string, TableRecord>();
            FilePath = filePath;
            if (offset > -1)
            {
                _reader.BaseStream.Position = offset;
            }
            if (format == FontFormat.Ttc)
            {
                InitializeTtc();
                return;
            }
            else
            {
                Initialize();
                ReadTableRecords();
            }
            if (_tableRecords.ContainsKey(TableNames.Name))
            {
                var tblSettings = new TableLoaderSettings(reader, _tableRecords, null);
                NameTable = TableLoaders.GetNameTableLoader(tblSettings).Load(false);                
                FontFamilyName = GetDefaultFontFamilyName();
                FontSubFamilyName = GetDefaultSubFontFamilyName();
                switch(FontSubFamilyName.ToLower())
                {
                    case "regular":
                        FontSubFamily = FontSubFamily.Regular;
                        break;
                    case "bold":
                        FontSubFamily = FontSubFamily.Bold;
                        break;
                    case "italic":
                        FontSubFamily = FontSubFamily.Italic;
                        break;
                    case "bold italic":
                        FontSubFamily = FontSubFamily.BoldItalic;
                        break;
                    default:
                        FontSubFamily = GetFontSubFamilyFromOs2(reader);
                        break;
                }
            }
        }

        private FontSubFamily GetFontSubFamilyFromOs2(FontsBinaryReader reader)
        {
            var tblSettings = new TableLoaderSettings(reader, _tableRecords, null);
            Os2Table = TableLoaders.GetOs2TableLoader(tblSettings).Load(false);
            if (EnumUtil.HasFlag(Os2Table.fsSelection, Os2Table.FsSelectionFlags.Bold | Os2Table.FsSelectionFlags.Italic))
            {
                return FontSubFamily.BoldItalic;
            }
            else if (EnumUtil.HasFlag(Os2Table.fsSelection, Os2Table.FsSelectionFlags.Bold))
            {
                return FontSubFamily.Bold;
            }
            else if (EnumUtil.HasFlag(Os2Table.fsSelection, Os2Table.FsSelectionFlags.Italic))
            {
                return FontSubFamily.Italic;
            }
            return FontSubFamily.Regular;
        }

        internal string GetDefaultFullFontFamilyName()
        {
            var v= NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping == null && x.RecordType == NameRecordTypes.FullFontName && x.LanguageMapping.Language == Languages.English)?.Name;
            if(v==null)
            {
                return GetEnglishFullFontFamilyName();
            }
            return v;
        }
        internal string GetEnglishFullFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FullFontName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        internal string GetDefaultFontFamilyName()
        {
            var v = NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping == null && x.RecordType == NameRecordTypes.FontFamilyName)?.Name;
            if (v == null)
            {
                return GetEnglishFontFamilyName();
            }
            return v;
        }
        public string GetEnglishFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontFamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        internal string GetDefaultSubFontFamilyName()
        {
            var v = NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping == null && x.RecordType == NameRecordTypes.FontSubfamilyName)?.Name;
            if (v == null)
            {
                return GetEnglishFontSubFamilyName();
            }
            return v;
        }
        internal string GetEnglishFontSubFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontSubfamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        private readonly FontsBinaryReader _reader;
        private ushort _numTables;
        private readonly Dictionary<string, TableRecord> _tableRecords;
        private FontFormat _format;

        public string FontFamilyName { get; private set; }

        public FontSubFamily FontSubFamily { get; set; }
        public string FontSubFamilyName { get; set; }

        public string FilePath { get; set; }

        public FontFormat Format { get; set; }

        public NameTable NameTable { get; private set; }
        public Os2Table Os2Table { get; private set; }

        public IEnumerable<ScannedFont>? SubFonts { get; private set; }

        public long? TtcOffset { get; set; }

        internal Dictionary<string, TableRecord> TableRecords => _tableRecords;


        private void InitializeTtc()
        {
            var tag = new Tag(_reader);
            var majorVersion = _reader.ReadUInt16BigEndian();
            var minorVersion = _reader.ReadUInt16BigEndian();
            var numFonts = _reader.ReadUInt32BigEndian();
            var offsets = new List<uint>();
            for (var x = 0; x < numFonts; x++)
            {
                offsets.Add(_reader.ReadUInt32BigEndian());
            }
            var fonts = new List<ScannedFont>();
            foreach (var offset in offsets)
            {
                var subFont = new ScannedFont(_reader, FontFormat.Ttf, FilePath, offset);
                subFont.TtcOffset = offset;
                fonts.Add(subFont);
            }
            SubFonts = fonts;
        }

        private void Initialize()
        {
            var sfntVersion = _reader.ReadUInt32BigEndian();
            // Number of tables.
            _numTables = _reader.ReadUInt16BigEndian();
            // Maximum power of 2 less than or equal to numTables,
            // times 16 ((2**floor(log2(numTables))) * 16,
            // where “**” is an exponentiation operator).
            var sr = _reader.ReadUInt16BigEndian();
            // Log2 of the maximum power of 2 less than or equal to
            // numTables (log2(searchRange/16), which is equal to
            // floor(log2(numTables))).
            var es = _reader.ReadUInt16BigEndian();
            // numTables times 16, minus searchRange
            // ((numTables * 16) - searchRange).
            var rs = _reader.ReadUInt16BigEndian();
        }

        private void ReadTableRecords()
        {
            for (var x = 0; x < _numTables; x++)
            {
                var record = new TableRecord
                {
                    Tag = new Tag(_reader),
                    Checksum = _reader.ReadUInt32BigEndian(),
                    Offset = _reader.ReadUInt32BigEndian(),
                    Length = _reader.ReadUInt32BigEndian()
                };
                _tableRecords.Add(record.Tag.Value, record);
            }
        }

        public byte[] GetTableBytes(string tag)
        {
            using var reader = new FontsBinaryReader(File.OpenRead(FilePath));
            if (!_tableRecords.TryGetValue(tag, out var record))
            {
                throw new ArgumentException($"Table '{tag}' not found in font.");
            }

            reader.BaseStream.Position = record.Offset;
            return reader.ReadBytes((int)record.Length);
        }

        //public override string ToString()
        //{
        //    if(Format == FontFormat.Ttc)
        //    {
        //        return $"Collection: {System.IO.Path.GetFileName(FilePath)}";
        //    }
        //    if(!string.IsNullOrEmpty(FontFamilyName))
        //    {
        //        return FontFamilyName.ToString();
        //    }
        //    return base.ToString();
        //}
    }
}
