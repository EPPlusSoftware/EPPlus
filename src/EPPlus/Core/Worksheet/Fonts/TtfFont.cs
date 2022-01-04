/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.FontLocalization;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Cmap;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Glyph;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Head;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Hhea;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Hmtx;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Kern;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Name;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Os2;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts
{
    internal class TtfFont
    {
        public TtfFont(BigEndianBinaryReader reader)
        {
            _reader = reader;
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
        }

        private readonly BigEndianBinaryReader _reader;

        private Dictionary<string, TableRecord> _tableRecords;

        

        //private uint[] GetGlyphIndexes(MyBinaryReader reader)
        //{
        //    var nGlyphs = GetNumberOfGlyphs(reader);
        //    var initialPos = reader.BaseStream.Position;
        //    var loader = new LocaTableLoader(reader, _tableRecords, nGlyphs);
        //    var table = loader.Load();
        //    reader.BaseStream.Position = initialPos;
        //    return table.Offsets.ToArray();
        //}

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

        public string GetEnglishFullFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FullFontName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        public string GetEnglishFontFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontFamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        public SerializedFontFamilies? GetFontFamily()
        {
            if(EnumCompatUtil.TryParse(GetEnglishFontFamilyName().Replace(" ", ""), out SerializedFontFamilies family))
            {
                return family;
            }
            return null;
        }

        public string GetEnglishFontSubFamilyName()
        {
            return NameTable.NameRecords.FirstOrDefault(x => x.LanguageMapping != null && x.RecordType == NameRecordTypes.FontSubfamilyName && x.LanguageMapping.Language == Languages.English)?.Name;
        }

        public FontSubFamilies? GetFontSubFamily()
        {
            if (EnumCompatUtil.TryParse(GetEnglishFontSubFamilyName().Replace(" ", ""), out FontSubFamilies subFamily))
            {
                return subFamily;
            }
            return null;
        }

        public uint SfntVersion { get; private set; }

        public ushort NumTables { get; private set; }

        public ushort SearchRange { get; private set; }

        public ushort EntrySelector { get; private set; }

        public ushort RangeShift { get; private set; }

        public IDictionary<string, TableRecord> TableRecords => _tableRecords;

        public CmapTable CmapTable { get; private set; }

        public NameTable NameTable { get; private set; }

        public GlyphTable GlyphTable { get; private set; }

        public Os2Table Os2Table { get; private set; }

        public HheaTable HheaTable { get; private set; }

        public HeadTable HeadTable { get; private set; }

        public HmtxTable HmtxTable { get; private set; }

        public KernTable KernTable { get; private set; }

    }
}
