using System;
using System.Collections.Generic;

namespace FontLab1.Tables.Os2
{
    internal class Os2TableLoader : TableLoader<Os2Table>
    {
        public Os2TableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Os2)
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
                panose.Add(BitConverter.ToInt16(new byte[] { p, 0 }, 0));
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

            var usWinAscent = _reader.ReadUInt16BigEndian();
            var usWinDescent = _reader.ReadUInt16BigEndian();
            var ulCodePageRange1 = _reader.ReadUInt32BigEndian();
            var ulCodePageRange2 = _reader.ReadUInt32BigEndian();
            var sxHeight = _reader.ReadInt16BigEndian();
            var sCapHeight = _reader.ReadInt16BigEndian();
            var usDefaultChar = _reader.ReadUInt16BigEndian();
            var usBreakChar = _reader.ReadUInt16BigEndian();
            var usMaxContext = _reader.ReadUInt16BigEndian();
            var usLowerOpticalPointSize = _reader.ReadUInt16BigEndian();
            var usUpperOpticalPointSize = _reader.ReadUInt16BigEndian();


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
                sTypoLineGap = sTypoLineGap,

                usWinAscent = usWinAscent,
                usWinDescent = usWinDescent,
                ulCodePageRange1 = ulCodePageRange1,
                ulCodePageRange2 = ulCodePageRange2,
                sxHeight = sxHeight,
                sCapHeight = sCapHeight,
                usDefaultChar = usDefaultChar,
                usBreakChar = usBreakChar,
                usMaxContext = usMaxContext,
                usLowerOpticalPointSize = usLowerOpticalPointSize,
                usUpperOpticalPointSize = usUpperOpticalPointSize,
            };
        }
    }
}
