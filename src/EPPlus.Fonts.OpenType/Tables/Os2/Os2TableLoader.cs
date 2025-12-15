using System;
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
namespace EPPlus.Fonts.OpenType.Tables.Os2
{
    internal class Os2TableLoader : TableLoader<Os2Table>
    {
        public Os2TableLoader(TableLoaderSettings settings) : base(settings, TableNames.Os2)
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
            var panose = new byte[10];
            for(var x = 0; x < 10; x++)
            {
                panose[x] = _reader.ReadByte();
            }
            var ucr1 = _reader.ReadUInt32BigEndian();
            var ucr2 = _reader.ReadUInt32BigEndian();
            var ucr3 = _reader.ReadUInt32BigEndian();
            var ucr4 = _reader.ReadUInt32BigEndian();
            var achVendId = new Tag(_reader);
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
            ushort usLowerOpticalPointSize = default;
            ushort usUpperOpticalPointSize = default;
            if (version > 3)
            {
                usLowerOpticalPointSize = _reader.ReadUInt16BigEndian();
                usUpperOpticalPointSize = _reader.ReadUInt16BigEndian();
            }
           


            var table = new Os2Table
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
                panose = panose,
                UnicodeRange1 = ucr1,
                UnicodeRange2 = ucr2,
                UnicodeRange3 = ucr3,
                UnicodeRange4 = ucr4,
                achVendId = achVendId,
                fsSelection = (Os2Table.FsSelectionFlags)fsSelection,
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
            };
            if(table.version > 3)
            {
                table.usLowerOpticalPointSize = usLowerOpticalPointSize;
                table.usUpperOpticalPointSize = usUpperOpticalPointSize;
            }
            return table;
        }
    }
}
