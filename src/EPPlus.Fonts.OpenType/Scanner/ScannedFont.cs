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
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Name;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Scanner
{
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
                FontFamilyName = NameTable.NameRecords.FirstOrDefault(x => x.RecordType == NameRecordTypes.FontFamilyName && !string.IsNullOrEmpty(x.Name))?.Name;
                FontSubFamilyName = NameTable.NameRecords.FirstOrDefault(x => x.RecordType == NameRecordTypes.FontSubfamilyName && !string.IsNullOrEmpty(x.Name))?.Name;
            }
        }

        private readonly FontsBinaryReader _reader;
        private ushort _numTables;
        private Dictionary<string, TableRecord> _tableRecords;
        private FontFormat _format;

        public string? FontFamilyName { get; private set; }

        public string? FontSubFamilyName { get; set; }

        public string FilePath { get; set; }

        public FontFormat Format { get; set; }

        public NameTable? NameTable { get; private set; }

        public IEnumerable<ScannedFont>? SubFonts { get; private set; }

        public long? TtcOffset { get; set; }


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
