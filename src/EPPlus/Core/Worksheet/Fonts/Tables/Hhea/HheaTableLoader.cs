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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Hhea
{
    internal class HheaTableLoader : TableLoader<HheaTable>
    {
        public HheaTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Hhea)
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
