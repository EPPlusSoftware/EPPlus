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
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Kern;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Base class for open-type fonts
    /// </summary>
    public class OpenTypeFont
    {
        internal TableCache _localTableCache;
        internal TableLoaderSettings _tblSettings;
        private readonly FontsBinaryReader _reader;
        protected Dictionary<string, TableRecord> _tableRecords;
        public FontFormat Format;
        private static object _syncRoot = new object();


        internal OpenTypeFont(FontFormat format)
        {
            Format = format;
            _tableRecords = new Dictionary<string, TableRecord>();
            _localTableCache = new TableCache();
        }


        internal OpenTypeFont(FontsBinaryReader reader, FontFormat format)
            : this(reader, -1, format)
        {
        }

        internal OpenTypeFont(FontsBinaryReader reader, long startOffset, FontFormat format)
        {
            Format = format;
            _reader = reader;

            lock (_syncRoot)
            {
                if (startOffset > -1)
                {
                    _reader.BaseStream.Position = startOffset;
                }

                Initialize();        // Reads SFNT header
                ReadTableRecords();  // Reads table directory
            }


            _localTableCache = new TableCache();

            _tblSettings = new TableLoaderSettings(_reader, _tableRecords, _localTableCache);

            //Ensure lazy-loading of individual tables via instanced table loaders.
            _os2TableLoader = TableLoaders.GetOs2TableLoader(_tblSettings);
            _nameTableLoader = TableLoaders.GetNameTableLoader(_tblSettings);
            _hheaTableLoader = TableLoaders.GetHheaTableLoader(_tblSettings);
            _headTableLoader = TableLoaders.GetHeadTableLoader(_tblSettings);
            _cmapTableLoader = TableLoaders.GetCmapTableLoader(_tblSettings);
            _hmtxTableLoader = TableLoaders.GetHmtxTableLoader(_tblSettings);
            _maxpTableLoader = TableLoaders.GetMaxpTableLoader(_tblSettings);
            _postTableLoader = TableLoaders.GetPostTableLoader(_tblSettings);
            _locaTableLoader = TableLoaders.GetLocaTableLoader(_tblSettings);
            _gsubTableLoader = TableLoaders.GetGsubTableLoader(_tblSettings);

            //Common tables in ttf fonts
            _glyfTableLoader = TableRecords.ContainsKey(TableNames.Glyf) ? TableLoaders.GetGlyfTableLoader(_tblSettings) : null;
            _kernTableLoader = TableRecords.ContainsKey(TableNames.Kern) ? TableLoaders.GetKernTableLoader(_tblSettings) : null;
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
        GsubTableLoader _gsubTableLoader;

        internal GlyfTableLoader _glyfTableLoader;
        internal KernTableLoader _kernTableLoader;

        internal FontSerializationContext GetSerializationContext()
        {
            return new FontSerializationContext(this);
        }

        public FontValidationReport ValidateFont(FontValidationSeverity severity)
        {
            var validator = new FontValidator();
            return validator.Validate(this);
        }


        internal List<uint> UsedCodePointsForSubset { get; set; } = new List<uint>();

        /// <summary>
        /// Any font file that does not contain all of the below tables can be considered corrupt as "the following tables are required for the font to function correctly"
        /// source: https://learn.microsoft.com/en-us/typography/opentype/spec/otff
        /// </summary>
        #region Required Font Tables 
        public CmapTable CmapTable
        {
            get
            {
                if (_cmapTableLoader != null)
                {
                    return _cmapTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Cmap))
                {
                    return (CmapTable)_localTableCache.Get(TableNames.Cmap);
                }
                return null;
            }
        }
        public HeadTable HeadTable
        {
            get
            {
                if (_headTableLoader != null)
                {
                    return _headTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Head))
                {
                    return (HeadTable)_localTableCache.Get(TableNames.Head);
                }
                return null;
            }
        }
        public HheaTable HheaTable
        {
            get
            {
                if (_hheaTableLoader != null)
                {
                    return _hheaTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Hhea))
                {
                    return (HheaTable)_localTableCache.Get(TableNames.Hhea);
                }
                return null;
            }
        }
        public HmtxTable HmtxTable
        {
            get
            {
                if (_hmtxTableLoader != null)
                {
                    return _hmtxTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Hmtx))
                {
                    return (HmtxTable)_localTableCache.Get(TableNames.Hmtx);
                }
                return null;
            }

        }
        public MaxpTable MaxpTable
        {
            get
            {
                if (_maxpTableLoader != null)
                {
                    return _maxpTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Maxp))
                {
                    return (MaxpTable)_localTableCache.Get(TableNames.Maxp);
                }
                return null;
            }
        }
        public NameTable NameTable
        {
            get
            {
                if (_nameTableLoader != null)
                {
                    return _nameTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Name))
                {
                    return (NameTable)_localTableCache.Get(TableNames.Name);
                }
                return null;
            }
        }
        public Os2Table Os2Table
        {
            get
            {
                if (_os2TableLoader != null)
                {
                    return _os2TableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Os2))
                {
                    return (Os2Table)_localTableCache.Get(TableNames.Os2);
                }
                return null;
            }
        }
        public PostTable PostTable
        {
            get
            {
                if (_postTableLoader != null)
                {
                    return _postTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Post))
                {
                    return (PostTable)_localTableCache.Get(TableNames.Post);
                }
                return null;
            }
        }

        public LocaTable LocaTable
        {
            get
            {
                if (_locaTableLoader != null)
                {
                    return _locaTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Loca))
                {
                    return (LocaTable)_localTableCache.Get(TableNames.Loca);
                }
                return null;
            }
        }
        #endregion

        //Extra accessors for common tables
        public GlyfTable GlyfTable
        {
            get
            {
                if (_glyfTableLoader != null)
                {
                    return _glyfTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Glyf))
                {
                    return (GlyfTable)_localTableCache.Get(TableNames.Glyf);
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
                if (_kernTableLoader != null)
                {
                    return _kernTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Kern))
                {
                    return (KernTable)_localTableCache.Get(TableNames.Kern);
                }
                else
                {
                    return null;
                }
            }
        }

        public GsubTable GsubTable
        {
            get
            {
                if (_gsubTableLoader != null)
                {
                    return _gsubTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Gsub))
                {
                    return (GsubTable)_localTableCache.Get(TableNames.Gsub);
                }
                else
                {
                    return null;
                }
            }
        }

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

        public string GetEnglishFullFontFamilyName()
        {
            return GetNameString(NameRecordTypes.FullFontName);
        }

        public string GetEnglishFontFamilyName()
        {
            return GetNameString(NameRecordTypes.FontFamilyName);
        }

        public string GetEnglishFontSubFamilyName()
        {
            return GetNameString(NameRecordTypes.FontSubfamilyName);
        }

        private string GetNameString(NameRecordTypes recordType)
        {
            // Priority 1: Windows English (Platform 3, Language 0x0409)
            var windowsEnglish = NameTable.NameRecords.FirstOrDefault(x =>
                x.platformId == 3 &&
                x.languageID == 0x0409 &&
                x.RecordType == recordType);

            if (windowsEnglish != null && !string.IsNullOrEmpty(windowsEnglish.Name))
                return windowsEnglish.Name;

            // Priority 2: Mac English (Platform 1, Language 0)
            var macEnglish = NameTable.NameRecords.FirstOrDefault(x =>
                x.platformId == 1 &&
                x.languageID == 0 &&
                x.RecordType == recordType);

            if (macEnglish != null && !string.IsNullOrEmpty(macEnglish.Name))
                return macEnglish.Name;

            // Priority 3: Unicode English (Platform 0, Language 0)
            var unicodeEnglish = NameTable.NameRecords.FirstOrDefault(x =>
                x.platformId == 0 &&
                x.languageID == 0 &&
                x.RecordType == recordType);

            if (unicodeEnglish != null && !string.IsNullOrEmpty(unicodeEnglish.Name))
                return unicodeEnglish.Name;

            // Priority 4: ANY Windows record of this type (fallback)
            var anyWindows = NameTable.NameRecords.FirstOrDefault(x =>
                x.platformId == 3 &&
                x.RecordType == recordType);

            if (anyWindows != null && !string.IsNullOrEmpty(anyWindows.Name))
                return anyWindows.Name;

            // Priority 5: ANY record of this type (last resort)
            var any = NameTable.NameRecords.FirstOrDefault(x =>
                x.RecordType == recordType);

            return any?.Name ?? string.Empty;
        }

        internal void AddOrReplaceTable<T>(T table)
            where T : FontTableBase
        {
            _localTableCache.AddOrReplace(table.Name, table);


            var record = new TableRecord
            {
                Tag = new Tag(table.Name),
                Length = (uint)table.GetLength(this),
                Offset = 0,
                Checksum = 0
            };

            if (_tableRecords.ContainsKey(table.Name))
            {
                _tableRecords.Remove(table.Name);
            }
            _tableRecords[table.Name] = record;

        }

        public OpenTypeFont CreateSubset(IEnumerable<char> usedChars)
        {

            var subsetBuilder = new SubsetFontBuilder();

            // Konvertera chars till Unicode code points
            var codePoints = usedChars
                .Select(c => (uint)c)      // tvinga unsigned
                .Distinct()
                .Where(cp => cp <= 0x10FFFF) // validering
                .Select(cp => (int)cp);

            // Skapa subset-font
            var newFont = subsetBuilder.CreateSubset(this, codePoints);

            var postProcessor = new SubsetPostProcessor();
            postProcessor.PostProcessSubset(newFont);

            return newFont;
        }

        public OpenTypeFont CreateSubset_Old(IEnumerable<char> usedChars)
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
            subsetFont.AddOrReplaceTable(HeadTable.Clone());
            subsetFont.AddOrReplaceTable(MaxpTable.Clone());
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

        internal Dictionary<string, byte[]> PreprocessedPaddedTables { get; } = new Dictionary<string, byte[]>();


        /// <summary>
        /// Total length (in bytes) of the underlying font stream.
        /// Returns 0 if reader is null.
        /// </summary>
        internal long FileLength
        {
            get
            {
                return _reader != null && _reader.BaseStream != null
                    ? _reader.BaseStream.Length
                    : 0L;
            }
        }

        public byte[] GetTableData(string tag)
        {
            if (_tableRecords.TryGetValue(tag, out var record))
            {
                if(_reader != null && record.Offset > 0)
                {
                    _reader.BaseStream.Position = record.Offset;
                    return _reader.ReadBytes((int)record.Length);
                }
            }
            var ctx = new FontSerializationContext(this);
            switch(tag)
            {
                case TableNames.Head:
                    return HeadTable.Serialize(ctx);
                case TableNames.Loca:
                    return LocaTable.Serialize(ctx);
                case TableNames.Cmap:
                    return CmapTable.Serialize(ctx);
                case TableNames.Glyf:
                    return GlyfTable.Serialize(ctx);
                case TableNames.Os2:
                    return Os2Table.Serialize(ctx);
                case TableNames.Hhea:
                    return HheaTable.Serialize(ctx);
                case TableNames.Maxp:
                    return MaxpTable.Serialize(ctx);
                case TableNames.Hmtx:
                    return HmtxTable.Serialize(ctx);
                case TableNames.Name:
                    return NameTable.Serialize(ctx);
                case TableNames.Kern:
                    return KernTable.Serialize(ctx);
                case TableNames.Post:
                    return PostTable.Serialize(ctx);
                case TableNames.Gsub:
                    return GsubTable.Serialize(ctx);
                default:
                    return null;
            }
        }

        public byte[] RawData
        {
            get
            {
                if (_reader != null && _reader.BaseStream != null)
                {
                    long originalPosition = _reader.BaseStream.Position;
                    try
                    {
                        _reader.BaseStream.Position = 0;
                        return _reader.ReadBytes((int)_reader.BaseStream.Length);
                    }
                    finally
                    {
                        _reader.BaseStream.Position = originalPosition;
                    }
                }
                return null;
            }
        }

        public byte[] Serialize()
        {
            var serializer = new OpenTypeFontSerializer(this);
            return serializer.Serialize();
        }
    }
}
