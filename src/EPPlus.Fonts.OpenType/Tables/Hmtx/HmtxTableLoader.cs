using System.Collections.Generic;

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
namespace EPPlus.Fonts.OpenType.Tables.Hmtx
{
    internal class HmtxTableLoader : TableLoader<HmtxTable>
    {
        private TableLoaderSettings tableLoaderSettingsRef;
        public HmtxTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Hmtx)
        {
            tableLoaderSettingsRef = settings;
        }

        protected override HmtxTable LoadInternal()
        {
            var hheaTable = TableLoaders.GetHheaTableLoader(tableLoaderSettingsRef).Load();
            var maxpTable = TableLoaders.GetMaxpTableLoader(tableLoaderSettingsRef).Load();
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
