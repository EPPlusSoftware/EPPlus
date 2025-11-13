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
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tables.Loca;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Base class for open-type fonts
    /// </summary>
    public class OpenTypeFont
    {
        internal TableCache localTableCache;
        internal TableLoaderSettings tblSettings;
        public FontFormat Format;

        internal OpenTypeFont(FontsBinaryReader reader, FontFormat format)
            : this(reader, -1, format)
        {
        }

        internal OpenTypeFont(FontsBinaryReader reader, long startOffset, FontFormat format)
        {
            Format = format;
            _reader = reader;
            if (startOffset > -1)
            {
                _reader.BaseStream.Position = startOffset;
            }
            Initialize();
            ReadTableRecords();

            localTableCache = new TableCache();

            tblSettings = new TableLoaderSettings(_reader, _tableRecords, localTableCache);

            //Ensure lazy-loading of individual tables via instanced table loaders.
            _os2TableLoader = TableLoaders.GetOs2TableLoader(tblSettings);
            _nameTableLoader = TableLoaders.GetNameTableLoader(tblSettings);
            _hheaTableLoader = TableLoaders.GetHheaTableLoader(tblSettings);
            _headTableLoader = TableLoaders.GetHeadTableLoader(tblSettings);
            _cmapTableLoader = TableLoaders.GetCmapTableLoader(tblSettings);
            _hmtxTableLoader = TableLoaders.GetHmtxTableLoader(tblSettings);
            _maxpTableLoader = TableLoaders.GetMaxpTableLoader(tblSettings);
            _postTableLoader = TableLoaders.GetPostTableLoader(tblSettings);
            _locaTableLoader = TableLoaders.GetLocaTableLoader(tblSettings);

            //Common tables in ttf fonts
            _glyfTableLoader = TableRecords.ContainsKey(TableNames.Glyf) ? TableLoaders.GetGlyfTableLoader(tblSettings) : null;
            _kernTableLoader = TableRecords.ContainsKey(TableNames.Kern) ? TableLoaders.GetKernTableLoader(tblSettings) : null;
        }

        Os2TableLoader _os2TableLoader;
        NameTableLoader _nameTableLoader;
        HheaTableLoader _hheaTableLoader;
        HeadTableLoader _headTableLoader;
        CmapTableLoader _cmapTableLoader;
        HmtxTableLoader _hmtxTableLoader;
        MaxpTableLoader _maxpTableLoader;
        PostTableLoader _postTableLoader;
        LocaTableLoader _locaTableLoader;

        internal GlyfTableLoader _glyfTableLoader;
        internal KernTableLoader _kernTableLoader;


        /// <summary>
        /// Any font file that does not contain all of the below tables can be considered corrupt as "the following tables are required for the font to function correctly"
        /// source: https://learn.microsoft.com/en-us/typography/opentype/spec/otff
        /// </summary>
        #region Required Font Tables 
        public CmapTable CmapTable 
        { 
            get 
            {
                return _cmapTableLoader.Load();
            } 
         }
        public HeadTable HeadTable
        {
            get
            {
                return _headTableLoader.Load();
            }
        }
        public HheaTable HheaTable
        {
            get
            {
                return _hheaTableLoader.Load();
            }
        }
        public HmtxTable HmtxTable
        {
            get
            {
                return _hmtxTableLoader.Load();
            }
        }
        public MaxpTable MaxpTable
        {
            get
            {
                return _maxpTableLoader.Load();
            }
        }
        public NameTable NameTable
        {
            get
            {
                return _nameTableLoader.Load();
            }
        }
        public Os2Table Os2Table
        {
            get
            {
                return _os2TableLoader.Load();
            }
        }
        public PostTable PostTable
        {
            get
            {
                return _postTableLoader.Load();
            }
        }

        public LocaTable LocaTable
        {
            get
            {
                return _locaTableLoader.Load();
            }
        }
        #endregion

        //Extra accessors for common tables
        public GlyfTable GlyfTable
        {
            get
            {
                if(_glyfTableLoader != null)
                {
                    return _glyfTableLoader.Load();
                }
                else
                {
                    return null;
                }
            }
        }
        public KernTable KernTable
        {
            get
            {
                if(_kernTableLoader != null)
                {
                    return _kernTableLoader.Load();
                }
                else
                {
                    return null;
                }
            }
        }
        private readonly FontsBinaryReader _reader;

        protected Dictionary<string, TableRecord> _tableRecords;

        private void Initialize()
        {
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

        public string FullName
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

        public string SubFamily
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

        public string GetEnglishFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontFamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        internal string GetEnglishFontSubFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontSubfamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        internal void ReplaceTable<T>(string tableName, T table)
            where T : FontTableBase
        {
            switch(tableName)
            {
                case TableNames.Head:
                    _headTableLoader.SetTable(TableNames.Head, table as HeadTable);
                    break;
                case TableNames.Maxp:
                    _maxpTableLoader.SetTable(TableNames.Maxp, table as MaxpTable); 
                    break;
                case TableNames.Glyf:
                    _glyfTableLoader.SetTable(TableNames.Glyf, table as GlyfTable);
                    break;
                case TableNames.Loca:
                    _locaTableLoader.SetTable(TableNames.Loca, table as LocaTable);
                    break;
                case TableNames.Hmtx:
                    _hmtxTableLoader.SetTable(TableNames.Hmtx, table as HmtxTable);
                    break;
                case TableNames.Cmap:
                    _cmapTableLoader.SetTable(TableNames.Cmap, table as CmapTable);
                    break;
                default:
                    return;
            }
        }

        public OpenTypeFont CreateSubset(IEnumerable<char> usedChars)
        {
            // 1. Map chars to glyph IDs
            var glyphIds = new HashSet<ushort>();
            foreach (var ch in usedChars)
            {
                var glyphId = CmapTable.MapCharToGlyph(ch);
                if (glyphId >= 0)
                    glyphIds.Add((ushort)glyphId);
            }
            glyphIds.Add(0); // Always include .notdef

            // 2. Handle composite glyphs
            GlyfTable.ResolveCompositeGlyphs(glyphIds);

            // 3. Create new font instance
            var subsetFont = new OpenTypeFont(_reader, Format);

            // 4. Copy and filter tables
            subsetFont.ReplaceTable(TableNames.Head, HeadTable.Clone());
            subsetFont.ReplaceTable(TableNames.Maxp, MaxpTable.Clone());
            subsetFont.MaxpTable.numGlyphs = (ushort)glyphIds.Count;

            //subsetFont.ReplaceTable(TableNames.Glyf, GlyfTable.CreateSubset(glyphIds));
            //subsetFont.LocaTable = this.LocaTable.CreateSubset(glyphIds);
            //subsetFont.HmtxTable = this.HmtxTable.CreateSubset(glyphIds);
            //subsetFont.CmapTable = this.CmapTable.CreateSubset(usedChars);

            //// 5. Recalculate checksums
            //subsetFont.RecalculateChecksums();

            return subsetFont;
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
                result = default;
                return false;
            }
        }

        internal uint SfntVersion { get; private set; }

        internal ushort NumTables { get; private set; }

        internal ushort SearchRange { get; private set; }

        internal ushort EntrySelector { get; private set; }

        internal ushort RangeShift { get; private set; }

        internal IDictionary<string, TableRecord> TableRecords => _tableRecords;
    }
}
