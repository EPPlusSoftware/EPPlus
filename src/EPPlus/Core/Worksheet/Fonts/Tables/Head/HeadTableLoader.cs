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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Head
{
    internal class HeadTableLoader : TableLoader<HeadTable>
    {
        public HeadTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Head)
        {
        }

        protected override HeadTable LoadInternal()
        {
            _reader.BaseStream.Position = _offset;
            var major = _reader.ReadUInt16BigEndian();
            var minor = _reader.ReadUInt16BigEndian();
            var fontRevision = _reader.ReadInt32();
            var checksumAdjustoment = _reader.ReadUInt32BigEndian();
            var magicNumber = _reader.ReadUInt32BigEndian();
            var flags = _reader.ReadUInt16BigEndian();
            var unitsPerEm = _reader.ReadUInt16BigEndian();
            var createdDate = _reader.ReadInt64();
            var modifiedDate = _reader.ReadInt64();
            var xMin = _reader.ReadInt16BigEndian();
            var yMin = _reader.ReadInt16BigEndian();
            var xMax = _reader.ReadInt16BigEndian();
            var yMax = _reader.ReadInt16BigEndian();
            var macStyle = _reader.ReadUInt16BigEndian();
            var lowestRecPPEM = _reader.ReadUInt16BigEndian();
            var fontDirectionHint = _reader.ReadInt16BigEndian();
            var indexToLocFormat = _reader.ReadInt16BigEndian();
            return new HeadTable
            {
                MajorVersion = major,
                MinorVersion = minor,
                UnitsPerEm = unitsPerEm,
                Xmin = xMin,
                Ymin = yMin,
                Xmax = xMax,
                Ymax = yMax,
                LowestRecPPEM = lowestRecPPEM,
                IndexToLocFormat = (HeadTable.IndexToLocFormats)indexToLocFormat
            };

        }
    }
}
