using System.Collections.Generic;

namespace FontLab1.Tables.Hhea
{
    internal class HheaTableLoader : TableLoader<HheaTable>
    {
        public HheaTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Hhea)
        {
        }

        protected override HheaTable LoadInternal()
        {
            var majorVersion = _reader.ReadUInt16BigEndian();
            var minorVersion = _reader.ReadUInt16BigEndian();
            var ascender = _reader.ReadInt16BigEndian();
            var descender = _reader.ReadInt16BigEndian();
            var lineGap = _reader.ReadInt16BigEndian();
            var advanceWidthMax = _reader.ReadUInt16BigEndian();
            var minLeftSideBearing = _reader.ReadInt16BigEndian();
            var minRightSideBearing = _reader.ReadInt16BigEndian();
            var xMaxExtent = _reader.ReadInt16BigEndian();
            var caretSlopeRise = _reader.ReadInt16BigEndian();
            var caretSlopeRun = _reader.ReadInt16BigEndian();
            var caretOffset = _reader.ReadInt16BigEndian();
            var reserved1 = _reader.ReadInt16BigEndian();
            var reserved2 = _reader.ReadInt16BigEndian();
            var reserved3 = _reader.ReadInt16BigEndian();
            var reserved4 = _reader.ReadInt16BigEndian();
            var metricDataFormat = _reader.ReadInt16BigEndian();
            var numberOfHMetrics = _reader.ReadUInt16BigEndian();

            return new HheaTable
            {
                majorVersion = majorVersion,
                minorVersion = minorVersion,
                ascender = ascender,
                descender = descender,
                lineGap = lineGap,
                advanceWidthMax = advanceWidthMax,
                minLeftSideBearing = minLeftSideBearing,
                minRightSideBearing = minRightSideBearing,
                xMaxExtent = xMaxExtent,
                caretSlopeRise = caretSlopeRise,
                caretSlopeRun = caretSlopeRun,
                caretOffset = caretOffset,
                metricDataFormat = metricDataFormat,
                numberOfHMetrics = numberOfHMetrics
            };
        }
    }
}
