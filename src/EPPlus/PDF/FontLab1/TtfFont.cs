using FontLab1.FontLocalization;
using FontLab1.Tables;
using FontLab1.Tables.Cmap;
using FontLab1.Tables.Glyph;
using FontLab1.Tables.Head;
using FontLab1.Tables.Hhea;
using FontLab1.Tables.Hmtx;
using FontLab1.Tables.Kern;
using FontLab1.Tables.Maxp;
using FontLab1.Tables.Name;
using FontLab1.Tables.Os2;
using FontLab1.Tables.Post;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using NameTable = FontLab1.Tables.Name.NameTable;

namespace FontLab1
{
    [DebuggerDisplay("{FullName} {SubFamily}")]
    internal class TtfFont
    {
        public TtfFont(MyBinaryReader reader)
            : this(reader, -1)
        {
        }

        public TtfFont(MyBinaryReader reader, long startOffset)
        {
            _reader = reader;
            if (startOffset > -1)
            {
                _reader.BaseStream.Position = startOffset;
            }
            Initialize();
            ReadTableRecords();
            Os2Table = TableLoaders.GetOs2TableLoader(reader, _tableRecords).Load();
            NameTable = TableLoaders.GetNameTableLoader(reader, _tableRecords).Load();
            HheaTable = TableLoaders.GetHheaTableLoader(reader, _tableRecords).Load();
            HeadTable = TableLoaders.GetHeadTableLoader(reader, _tableRecords).Load();
            CmapTable = TableLoaders.GetCmapTableLoader(reader, _tableRecords).Load();
            GlyphTable = TableLoaders.GetGlyphTableLoader(reader, _tableRecords).Load();
            HmtxTable = TableLoaders.GetHtmxTableLoader(reader, _tableRecords).Load();
            KernTable = TableLoaders.GetKernTableLoader(reader, _tableRecords).Load();
            postTable = TableLoaders.GetPostTableLoader(reader, _tableRecords).Load();
        }

        private readonly MyBinaryReader _reader;

        private Dictionary<string, TableRecord> _tableRecords;

        private void Initialize()
        {
            TableCache.Clear();

            SfntVersion = _reader.ReadUInt32BigEndian();
            // Number of tables.
            NumTables = _reader.ReadUInt16BigEndian();
            // Maximum power of 2 less than or equal to numTables,
            // times 16 ((2**floor(log2(numTables))) * 16,
            // where “**” is an exponentiation operator).
            SearchRange = _reader.ReadUInt16BigEndian();
            // Log2 of the maximum power of 2 less than or equal to
            // numTables (log2(searchRange/16), which is equal to
            // floor(log2(numTables))).
            EntrySelector = _reader.ReadUInt16BigEndian();
            // numTables times 16, minus searchRange
            // ((numTables * 16) - searchRange).
            RangeShift = _reader.ReadUInt16BigEndian();
        }

        private void ReadTableRecords()
        {
            _tableRecords = new Dictionary<string, TableRecord>();
            for (var x = 0; x < NumTables; x++)
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

        internal string FullName
        {
            get
            {
                var n = GetEnglishFullFontFamilyName();
                if (string.IsNullOrEmpty(n))
                {
                    return "Unknown font";
                }
                return n;
            }
        }

        internal string SubFamily
        {
            get
            {
                var n = GetEnglishFontSubFamilyName();
                if (string.IsNullOrEmpty(n))
                {
                    return "Unknown subfamily";
                }
                return n;
            }
        }

        internal string GetEnglishFullFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FullFontName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        internal string GetEnglishFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontFamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        internal string GetEnglishFontSubFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontSubfamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        public static bool TryParseEnum<T>(string value, out T result) where T : struct
        {
            try
            {
                result = (T)Enum.Parse(typeof(T), value, ignoreCase: true);
                return true;
            }
            catch
            {
                result = default(T);
                return false;
            }
        }

        internal uint SfntVersion { get; private set; }

        internal ushort NumTables { get; private set; }

        internal ushort SearchRange { get; private set; }

        internal ushort EntrySelector { get; private set; }

        internal ushort RangeShift { get; private set; }

        internal IDictionary<string, TableRecord> TableRecords => _tableRecords;

        internal CmapTable CmapTable { get; private set; }

        internal NameTable NameTable { get; private set; }

        internal GlyphTable GlyphTable { get; private set; }

        internal Os2Table Os2Table { get; private set; }

        internal HheaTable HheaTable { get; private set; }

        internal HeadTable HeadTable { get; private set; }

        internal HmtxTable HmtxTable { get; private set; }

        internal KernTable KernTable { get; private set; }

        internal PostTable postTable { get; private set; }

        internal MaxpTable maxpTable { get; private set; }

    }
}
