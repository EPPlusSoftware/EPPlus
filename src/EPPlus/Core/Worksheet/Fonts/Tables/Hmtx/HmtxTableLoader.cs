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
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Hmtx
{
    public class HmtxTableLoader : TableLoader<HmtxTable>
    {
        public HmtxTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Hmtx)
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
