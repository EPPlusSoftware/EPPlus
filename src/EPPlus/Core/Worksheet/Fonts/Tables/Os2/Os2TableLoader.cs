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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Os2
{
    public class Os2TableLoader : TableLoader<Os2Table>
    {
        public Os2TableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Os2)
        {
        }

        protected override Os2Table LoadInternal()
        {
            var version = _reader.ReadUInt16BigEndian();
            var xAvgCharWidth = _reader.ReadInt16BigEndian();
            var usWeightClass = _reader.ReadUInt16BigEndian();
            var usWidthClass = _reader.ReadUInt16BigEndian();
            var fsType = _reader.ReadUInt16BigEndian();
            var ySubscriptXSize = _reader.ReadInt16BigEndian();
            var ySubscriptYSize = _reader.ReadInt16BigEndian();
            var ySubscriptXOffset = _reader.ReadInt16BigEndian();
            var ySubscriptYOffset = _reader.ReadInt16BigEndian();
            var ySuperscriptXSize = _reader.ReadInt16BigEndian();
            var ySuperscriptYSize = _reader.ReadInt16BigEndian();
            var ySuperscriptXOffset = _reader.ReadInt16BigEndian();
            var ySuperscriptYOffset = _reader.ReadInt16BigEndian();
            var yStrikeoutSize = _reader.ReadInt16BigEndian();
            var yStrikeoutPosition = _reader.ReadInt16BigEndian();
            var familyClass = _reader.ReadInt16BigEndian();
            // read panose
            var panose = new List<short>();
            for(var x = 0; x < 10; x++)
            {
                var p = _reader.ReadByte();
                panose.Add(BitConverter.ToInt16(new byte[] { p, 0 }));
            }
            var ucr1 = _reader.ReadUInt32BigEndian();
            var ucr2 = _reader.ReadUInt32BigEndian();
            var ucr3 = _reader.ReadUInt32BigEndian();
            var ucr4 = _reader.ReadUInt32BigEndian();
            var archVendId = new Tag(_reader);
            var fsSelection = _reader.ReadUInt16BigEndian();
            var usFirstCharIndex = _reader.ReadUInt16BigEndian();
            var usLastCharIndex = _reader.ReadUInt16BigEndian();
            var sTypoAscender = _reader.ReadInt16BigEndian();
            var sTypoDescender = _reader.ReadInt16BigEndian();
            var sTypoLineGap = _reader.ReadInt16BigEndian();

            return new Os2Table
            {
                version = version,
                xAvgCharWidth = xAvgCharWidth,
                usWeightClass = usWeightClass,
                usWidthClass = usWidthClass,
                fsType = fsType,
                ySubscriptXSize = ySubscriptXSize,
                ySubscriptYSize = ySubscriptYSize,
                ySubscriptXOffset = ySubscriptXOffset,
                ySubscriptYOffset = ySubscriptYOffset,
                ySuperscriptXSize = ySuperscriptXSize,
                ySuperscriptYSize = ySuperscriptYSize,
                ySuperscriptXOffset = ySuperscriptXOffset,
                ySuperscriptYOffset = ySuperscriptYOffset,
                yStrikeoutSize = yStrikeoutSize,
                yStrikeoutPosition = yStrikeoutPosition,
                sFamilyClass = familyClass,
                panose = panose.ToArray(),
                UnicodeRange1 = ucr1,
                UnicodeRange2 = ucr2,
                UnicodeRange3 = ucr3,
                UnicodeRange4 = ucr4,
                archVendId = archVendId,
                fsSelection = fsSelection,
                usFirstCharIndex = usFirstCharIndex,
                usLastCharIndex = usLastCharIndex,
                sTypoAscender = sTypoAscender,
                sTypoDescender = sTypoDescender,
                sTypoLineGap = sTypoLineGap
            };
        }
    }
}
