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
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Gpos;
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
using EPPlus.Fonts.OpenType.Tables.Vhea;
using EPPlus.Fonts.OpenType.Tables.Vmtx;
using EPPlus.Fonts.OpenType.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Base class for open-type fonts
    /// </summary>
    [DebuggerDisplay("{FullName}, IsSubset: {IsSubset}")]
    public class OpenTypeFont
    {
        internal TableCache _localTableCache;
        internal TableLoaderCache _loaderCache;
        internal TableLoaderSettings _tblSettings;
        private readonly byte[] _fontBytes;
        protected Dictionary<string, TableRecord> _tableRecords;
        public FontFormat Format;
        private readonly object _syncRoot = new object();


        internal OpenTypeFont(FontFormat format, bool isSubset = false)
        {
            Format = format;
            _tableRecords = new Dictionary<string, TableRecord>();
            _localTableCache = new TableCache();
            _loaderCache = new TableLoaderCache();
            IsSubset = isSubset;
        }


        internal OpenTypeFont(byte[] fontBytes)
            : this(fontBytes, -1)
        {
        }

        internal OpenTypeFont(byte[] fontBytes, long startOffset)
        {
            if (fontBytes == null || fontBytes.Length < 4)
                throw new ArgumentException("Invalid font data: too short to contain a valid SFNT header.", nameof(fontBytes));

            Format = DetectFormat(fontBytes, startOffset > -1 ? startOffset : 0);

            _fontBytes = fontBytes;
            var tableReaderFactory = new FontTableReaderFactory(fontBytes);
            using var reader = tableReaderFactory.CreateReader(startOffset);

            lock (_syncRoot)
            {
                if (startOffset > -1)
                {
                    reader.BaseStream.Position = startOffset;
                }

                Initialize(reader);        // Reads SFNT header
                ReadTableRecords(reader);  // Reads table directory
            }


            _localTableCache = new TableCache();
            _loaderCache = new TableLoaderCache();
            _tblSettings = new TableLoaderSettings(tableReaderFactory, _tableRecords, _localTableCache, _loaderCache);

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

            // ✅ Optional tables - only create loader if table exists
            _gsubTableLoader = TableRecords.ContainsKey(TableNames.Gsub)
                ? TableLoaders.GetGsubTableLoader(_tblSettings)
                : null;
            _gposTableLoader = TableRecords.ContainsKey(TableNames.Gpos)
                ? TableLoaders.GetGposTableLoader(_tblSettings)
                : null;
            _glyfTableLoader = TableRecords.ContainsKey(TableNames.Glyf)
                ? TableLoaders.GetGlyfTableLoader(_tblSettings)
                : null;
            _kernTableLoader = TableRecords.ContainsKey(TableNames.Kern)
                ? TableLoaders.GetKernTableLoader(_tblSettings)
                : null;
            _vheaTableLoader = TableRecords.ContainsKey(TableNames.Vhea)
                ? TableLoaders.GetVheaTableLoader(_tblSettings)
                : null;
            _vmtxTableLoader = TableRecords.ContainsKey(TableNames.Vmtx)
               ? TableLoaders.GetVmtxTableLoader(_tblSettings)
               : null;
        }

        /// <summary>
        /// Detects the font format from the SFNT version field in the header.
        /// </summary>
        /// <param name="fontBytes">Raw font bytes</param>
        /// <param name="offset">Offset to the start of the SFNT header</param>
        /// <returns>Detected FontFormat</returns>
        /// <exception cref="ArgumentException">Thrown if the header contains an unrecognized SFNT version</exception>
        private static FontFormat DetectFormat(byte[] fontBytes, long offset)
        {
            // sfntVersion is a big-endian UInt32 at the start of the SFNT header
            uint sfntVersion =
                ((uint)fontBytes[offset + 0] << 24) |
                ((uint)fontBytes[offset + 1] << 16) |
                ((uint)fontBytes[offset + 2] << 8) |
                ((uint)fontBytes[offset + 3]);

            switch (sfntVersion)
            {
                case 0x00010000: // TrueType
                case 0x74727565: // 'true' — Apple TrueType
                    return FontFormat.Ttf;

                case 0x4F54544F: // 'OTTO' — OpenType/CFF
                case 0x74797031: // 'typ1' — PostScript Type 1
                    return FontFormat.Otf;

                default:
                    throw new ArgumentException(
                        $"Unrecognized SFNT version 0x{sfntVersion:X8}. " +
                        "The data does not appear to be a valid TTF or OTF font.",
                        "fontBytes");
            }
        }

        Os2TableLoader _os2TableLoader;
        NameTableLoader _nameTableLoader;
        HheaTableLoader _hheaTableLoader;
        VheaTableLoader _vheaTableLoader;
        HeadTableLoader _headTableLoader;
        CmapTableLoader _cmapTableLoader;
        HmtxTableLoader _hmtxTableLoader;
        VmtxTableLoader _vmtxTableLoader;
        MaxpTableLoader _maxpTableLoader;
        PostTableLoader _postTableLoader;
        LocaTableLoader _locaTableLoader;
        GsubTableLoader _gsubTableLoader;
        GposTableLoader _gposTableLoader;

        internal GlyfTableLoader _glyfTableLoader;
        internal KernTableLoader _kernTableLoader;
        private volatile bool _fullyLoaded = false;

        public bool FullyLoaded => _fullyLoaded;
        internal bool IsReadOnly { get; set; }

        internal void EnsureFullyLoaded()
        {
            if (_fullyLoaded)
                return;

            lock (_loaderCache.SyncLock)
            {

                // --- Required tables (always present in valid fonts) ---
                // Each property accessor calls its TableLoader.Load() which
                // reads from the byte[] stream and caches the result.
                // By accessing them all here under the font-level lock,
                // we guarantee no concurrent reader access.
                FontTableBase _ = CmapTable;
                _ = HeadTable;
                _ = HheaTable;
                _ = HmtxTable;
                _ = MaxpTable;
                _ = NameTable;
                _ = Os2Table;
                _ = PostTable;
                _ = LocaTable;

                // --- Optional tables (only if present in font) ---
                if (_gsubTableLoader != null)
                {
                    var gsub = GsubTable;  // Forces full GSUB parse
                }
                if (_gposTableLoader != null)
                {
                    var gpos = GposTable;  // Forces full GPOS parse (incl. MarkToBase subtables)
                }
                if (_glyfTableLoader != null)
                {
                    var glyf = GlyfTable;  // Forces full glyph outline parse
                }
                if (_kernTableLoader != null)
                {
                    var kern = KernTable;  // Forces legacy kern table parse
                }

                _fullyLoaded = true;
            }
        }

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

        public VheaTable VheaTable
        {
            get
            {
                if (_vheaTableLoader != null)
                {
                    return _vheaTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Vhea))
                {
                    return (VheaTable)_localTableCache.Get(TableNames.Vhea);
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

        public VmtxTable VmtxTable
        {
            get
            {
                if (_vmtxTableLoader != null)
                {
                    return _vmtxTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Vmtx))
                {
                    return (VmtxTable)_localTableCache.Get(TableNames.Vmtx);
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

        public GposTable GposTable
        {
            get
            {
                if (_gposTableLoader != null)
                {
                    return _gposTableLoader.Load();
                }
                else if (_localTableCache.Contains(TableNames.Gpos))
                {
                    return (GposTable)_localTableCache.Get(TableNames.Gpos);
                }
                else
                {
                    return null;
                }
            }
        }

        private void Initialize(FontsBinaryReader reader)
        {
            SfntVersion = reader.ReadUInt32BigEndian();
            // Number of tables.
            NumTables = reader.ReadUInt16BigEndian();
            // Maximum power of 2 less than or equal to numTables,
            // times 16 ((2**floor(log2(numTables))) * 16,
            // where “**” is an exponentiation operator).
            SearchRange = reader.ReadUInt16BigEndian();
            // Log2 of the maximum power of 2 less than or equal to
            // numTables (log2(searchRange/16), which is equal to
            // floor(log2(numTables))).
            EntrySelector = reader.ReadUInt16BigEndian();
            // numTables times 16, minus searchRange
            // ((numTables * 16) - searchRange).
            RangeShift = reader.ReadUInt16BigEndian();
        }

        private void ReadTableRecords(FontsBinaryReader reader)
        {
            _tableRecords = new Dictionary<string, TableRecord>();
            for (var x = 0; x < NumTables; x++)
            {
                var record = new TableRecord
                {
                    Tag = new Tag(reader),
                    Checksum = reader.ReadUInt32BigEndian(),
                    Offset = reader.ReadUInt32BigEndian(),
                    Length = reader.ReadUInt32BigEndian()
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

        public bool IsSubset
        {
            get; private set;
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
            if (IsReadOnly)
                throw new InvalidOperationException(
                    $"Cannot modify a cached font instance. Table: {table.Name}. " +
                    "Use CreateSubset() or create a new OpenTypeFont instance.");
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
            // Validate input
            if (usedChars == null)
                throw new ArgumentNullException(nameof(usedChars));

            if (!usedChars.Any())
                throw new ArgumentException("Text cannot be empty", nameof(usedChars));

            var subsetBuilder = new SubsetFontBuilder();

            // Extract Unicode code points, correctly handling surrogate pairs.
            // A string like "Hello 😀" contains 7 chars but 6 code points,
            // because 😀 (U+1F600) is encoded as two UTF-16 surrogates.
            var codePoints = CodePointUtil.ExtractCodePoints(usedChars);

            var newFont = subsetBuilder.CreateSubset(this, codePoints);

            var postProcessor = new SubsetPostProcessor();
            postProcessor.PostProcessSubset(newFont);

            return newFont;
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
        /// For subset fonts: Maps original glyph IDs to new subset glyph IDs.
        /// Null for non-subset fonts.
        /// </summary>
        public Dictionary<ushort, ushort> SubsetGlyphMapping { get; internal set; }


        /// <summary>
        /// Total length (in bytes) of the underlying font stream.
        /// Returns 0 if reader is null.
        /// </summary>
        internal long FileLength
        {
            get
            {
                if(_tblSettings == null || _tblSettings.TableReaderFactory == null)
                {
                    return 0L;
                }
                return _tblSettings.TableReaderFactory.FontBytesLength;
            }
        }

        public byte[] GetTableData(string tag)
        {
            if (_tableRecords.TryGetValue(tag, out var record) && _tblSettings != null && _tblSettings.TableReaderFactory != null)
            {
                using var reader = _tblSettings.TableReaderFactory.CreateReader();
                if (reader != null && record.Offset > 0)
                {
                    reader.BaseStream.Position = record.Offset;
                    return reader.ReadBytes((int)record.Length);
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
                case TableNames.Gpos:
                    return GposTable.Serialize(ctx);
                case TableNames.Vhea:
                    return VheaTable?.Serialize(ctx);
                case TableNames.Vmtx:
                    return VmtxTable?.Serialize(ctx);
                default:
                    return null;
            }
        }

        public byte[] RawData
        {
            get
            {
                var reader = default(FontsBinaryReader);
                if(_tblSettings != null && _tblSettings.TableReaderFactory != null)
                {
                    reader = _tblSettings.TableReaderFactory.CreateReader();
                }
                if (reader != null && reader.BaseStream != null)
                {
                    long originalPosition = reader.BaseStream.Position;
                    try
                    {
                        reader.BaseStream.Position = 0;
                        return reader.ReadBytes((int)reader.BaseStream.Length);
                    }
                    finally
                    {
                        reader.BaseStream.Position = originalPosition;
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
