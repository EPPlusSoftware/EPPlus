using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Head;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Loca
{
    internal class LocaTableLoader : TableLoader<LocaTable>
    {
        public LocaTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Loca)
        {

        }

        protected override LocaTable LoadInternal()
        {
            var headTable = TableLoaders.GetHeadTableLoader(_reader, _tables).Load();
            var maxpTable = TableLoaders.GetMaxpTableLoader(_reader, _tables).Load();
            _reader.BaseStream.Position = _offset;
            var indexes = new List<uint>();
            if(headTable.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                for(var x = 0; x <= maxpTable.numGlyphs + 1; x++)
                {
                    var ix = Convert.ToUInt32(_reader.ReadUInt16BigEndian());
                    ix *= 2;
                    indexes.Add(ix);
                }
            }
            else if(headTable.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset32)
            {
                for(var x = 0; x <= maxpTable.numGlyphs + 1; x++)
                {
                    var ix = _reader.ReadUInt32BigEndian();
                    indexes.Add(ix);
                }
                
            }
            return new LocaTable
            {
                Offsets = indexes.ToArray()
            };
        }
    }
}
