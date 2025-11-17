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
        internal TableCache _localTableCache;
        internal TableLoaderSettings _tblSettings;
        private readonly FontsBinaryReader _reader;
        protected Dictionary<string, TableRecord> _tableRecords;
        public FontFormat Format;


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
            if (startOffset > -1)
            {
                _reader.BaseStream.Position = startOffset;
            }
            Initialize();
            ReadTableRecords();

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
                if(_cmapTableLoader != null)
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
                if(_headTableLoader != null)
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
                if(_hheaTableLoader != null)
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
                if(_nameTableLoader != null)
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
                if(_os2TableLoader != null)
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
                if(_postTableLoader != null)
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
                if(_locaTableLoader != null)
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
                if(_glyfTableLoader != null)
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
                if(_kernTableLoader != null)
                {
                    return _kernTableLoader.Load();
                }
                else if(_localTableCache.Contains(TableNames.Kern))
                {
                    return (KernTable)_localTableCache.Get(TableNames.Kern);
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

        internal T AddOrReplaceTable<T>(string tableName, T table)
            where T : FontTableBase
        {
            _localTableCache.AddOrReplace(tableName, table);


            var record = new TableRecord
            {
                Tag = new Tag(tableName),
                Length = (uint)table.GetLength(),
                Offset = 0,
                Checksum = 0
            };

            if(_tableRecords.ContainsKey(tableName))
            {
                _tableRecords.Remove(tableName);
            }
            _tableRecords[tableName] = record;
            return table;
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

            // 3. Sort glyph IDs and create mapping old -> new
            var sortedGlyphIds = glyphIds.OrderBy(id => id).ToList();
            var idMapping = new Dictionary<ushort, ushort>();
            for (ushort i = 0; i < sortedGlyphIds.Count; i++)
            {
                idMapping[sortedGlyphIds[i]] = i;
            }


            // 4. Create new font instance
            var subsetFont = new OpenTypeFont(_reader, Format);

            // 5. Copy and filter tables
            subsetFont.AddOrReplaceTable(TableNames.Head, HeadTable.Clone());
            subsetFont.AddOrReplaceTable(TableNames.Maxp, MaxpTable.Clone());
            subsetFont.MaxpTable.numGlyphs = (ushort)glyphIds.Count;
            var newGlyf = subsetFont.AddOrReplaceTable(TableNames.Glyf, GlyfTable.CreateSubset(sortedGlyphIds, idMapping));
            subsetFont.AddOrReplaceTable(TableNames.Loca, newGlyf);
            subsetFont.AddOrReplaceTable(TableNames.Hmtx, HmtxTable.CreateSubset(glyphIds, idMapping));
            subsetFont.AddOrReplaceTable(TableNames.Cmap, CmapTable.CreateSubset(usedChars, idMapping));

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

        public void RecalculateChecksums()
        {
            uint totalSum = 0;

            // Calculate checksums for each table
            foreach (var entry in TableRecords)
            {
                byte[] tableData = entry.Value.GetTableBytes(this);
                uint checksum = CalculateChecksum(tableData);
                entry.Value.Length = (uint)tableData.Length;
                entry.Value.Checksum = checksum;
                totalSum += checksum;
            }

            // Add sum of table directory entries
            foreach (var dirEntry in TableRecords)
            {
                totalSum += (uint)dirEntry.Value.Tag.Bytes[0] << 24 |
                            (uint)dirEntry.Value.Tag.Bytes[1] << 16 |
                            (uint)dirEntry.Value.Tag.Bytes[2] << 8 |
                            (uint)dirEntry.Value.Tag.Bytes[3];
                totalSum += dirEntry.Value.Checksum;
                totalSum += dirEntry.Value.Offset;
                totalSum += dirEntry.Value.Length;
            }

            // Adjust head.checkSumAdjustment
            uint adjustment = 0xB1B0AFBA - totalSum;
            HeadTable.ChecksumAdjustment = adjustment;
        }

        private uint CalculateChecksum(byte[] data)
        {
            uint sum = 0;
            int length = data.Length;
            int i = 0;

            while (length > 3)
            {
                sum += (uint)(data[i] << 24 | data[i + 1] << 16 | data[i + 2] << 8 | data[i + 3]);
                i += 4;
                length -= 4;
            }

            // Pad remaining bytes
            uint last = 0;
            for (int j = 0; j < length; j++)
            {
                last |= (uint)data[i + j] << (24 - j * 8);
            }
            sum += last;

            return sum;
        }
    }
}
