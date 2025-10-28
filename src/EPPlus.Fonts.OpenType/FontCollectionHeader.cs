using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Utils.Platform;
using System;
using System.Collections.Generic;
using EPPlus.Fonts.OpenType.Tables;

namespace EPPlus.Fonts.OpenType
{
    enum FontCollectionType
    {
        version1,
        version2
    }

    internal class FontCollectionHeader
    {
        FontsBinaryReader _reader;
        Tag _tag;
        ushort _majorVersion;
        ushort _minorVersion;
        uint _numFonts;


        public FontCollectionHeader(FontsBinaryReader reader) : this(reader, -1)
        {
        }

        public FontCollectionHeader(FontsBinaryReader reader, long startOffset)
        {
            _reader = reader;
            if (startOffset > -1)
            {
                _reader.BaseStream.Position = startOffset;
            }
            InitializeTtc();

            //ReadTableRecords();
            //Os2Table = TableLoaders.GetOs2TableLoader(reader, _tableRecords).Load();
            //NameTable = TableLoaders.GetNameTableLoader(reader, _tableRecords).Load();
            //HheaTable = TableLoaders.GetHheaTableLoader(reader, _tableRecords).Load();
            //HeadTable = TableLoaders.GetHeadTableLoader(reader, _tableRecords).Load();
            //CmapTable = TableLoaders.GetCmapTableLoader(reader, _tableRecords).Load();
            //HmtxTable = TableLoaders.GetHtmxTableLoader(reader, _tableRecords).Load();
            //postTable = TableLoaders.GetPostTableLoader(reader, _tableRecords).Load();
        }

        private void InitializeTtc()
        {
            _tag = new Tag(_reader);
            _majorVersion = _reader.ReadUInt16BigEndian();
            _minorVersion = _reader.ReadUInt16BigEndian();
            _numFonts = _reader.ReadUInt32BigEndian();
            var offsets = new List<uint>();
            for (var x = 0; x < _numFonts; x++)
            {
                offsets.Add(_reader.ReadUInt32BigEndian());
            }
        }
    }
}
