using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Hmtx
{
    internal class HmtxTableLoader : TableLoader<HmtxTable>
    {
        public HmtxTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Hmtx)
        {

        }


        protected override HmtxTable LoadInternal()
        {
            var hheaTable = TableLoaders.GetHheaTableLoader(_reader, _tables).Load();
            var maxpTable = TableLoaders.GetMaxpTableLoader(_reader, _tables).Load();
            _reader.BaseStream.Position = _offset;
            var metrics = new List<LongHorMetric>();
            for(var x = 0; x  < hheaTable.numberOfHMetrics; x++)
            {
                var metric = new LongHorMetric
                {
                    advanceWidth = _reader.ReadUInt16BigEndian(),
                    lsb = _reader.ReadInt16BigEndian()
                };
                metrics.Add(metric);
            }
            var bearings = new List<short>();
            for(var x = 0; x < (maxpTable.numGlyphs - hheaTable.numberOfHMetrics); x++)
            {
                var b = _reader.ReadInt16BigEndian();
                bearings.Add(b);
            }
            return new HmtxTable
            {
                hMetrics = metrics.ToArray(),
                leftSideBearings = bearings.ToArray()
            };
        }
    }
}
