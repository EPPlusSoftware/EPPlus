/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/18/2026         EPPlus Software AB           vhea table implementation (vertical text support)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Vhea;

namespace EPPlus.Fonts.OpenType.Tables
{
    internal class VheaTableLoader : TableLoader<VheaTable>
    {
        public VheaTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Vhea)
        {
        }

        protected override VheaTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;

            var table = new VheaTable();

            // Version is a Fixed (uint32): 0x00011000 = v1.1, 0x00010000 = v1.0
            table.Version = _reader.ReadUInt32BigEndian();
            table.Ascent = _reader.ReadInt16BigEndian();
            table.Descent = _reader.ReadInt16BigEndian();
            table.LineGap = _reader.ReadInt16BigEndian();
            table.AdvanceHeightMax = _reader.ReadInt16BigEndian();
            table.MinTopSideBearing = _reader.ReadInt16BigEndian();
            table.MinBottomSideBearing = _reader.ReadInt16BigEndian();
            table.YMaxExtent = _reader.ReadInt16BigEndian();
            table.CaretSlopeRise = _reader.ReadInt16BigEndian();
            table.CaretSlopeRun = _reader.ReadInt16BigEndian();
            table.CaretOffset = _reader.ReadInt16BigEndian();
            table.Reserved1 = _reader.ReadInt16BigEndian();
            table.Reserved2 = _reader.ReadInt16BigEndian();
            table.Reserved3 = _reader.ReadInt16BigEndian();
            table.Reserved4 = _reader.ReadInt16BigEndian();
            table.MetricDataFormat = _reader.ReadInt16BigEndian();
            table.NumberOfVMetrics = _reader.ReadUInt16BigEndian();

            return table;
        }
    }
}