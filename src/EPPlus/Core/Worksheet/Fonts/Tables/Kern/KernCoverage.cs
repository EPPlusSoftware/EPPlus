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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Kern
{
    public class KernCoverage
    {
        public KernCoverage(BigEndianBinaryReader reader)
        {
            _coverage = reader.ReadUInt16BigEndian();
        }

        public ushort _coverage;

        /// <summary>
        /// true if table has horizontal data, false if vertical.
        /// </summary>
        public bool horizontal
        {
            get
            {
                return (_coverage & 0x1) == 1;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool minimum
        {
            get
            {
                return ((_coverage >> 1) & 0x1) == 1;
            }
        }

        /// <summary>
        /// If set to 1, kerning is perpendicular to the flow of the text.
        /// 
        /// If the text is normally written horizontally, kerning will be done 
        /// in the up and down directions. If kerning values are positive, the 
        /// text will be kerned upwards; if they are negative, the text will be 
        /// kerned downwards.
        /// 
        /// If the text is normally written vertically, kerning will be done in 
        /// the left and right directions. If kerning values are positive, the 
        /// text will be kerned to the right; if they are negative, the text will 
        /// be kerned to the left.
        /// 
        /// The value 0x8000 in the kerning data resets the cross-stream kerning back to 0.
        /// </summary>
        public bool crossStream
        {
            get
            {
                return ((_coverage >> 2) & 0x1) == 1;
            }
        }

        /// <summary>
        /// If true the value in this table should replace the value 
        /// currently being accumulated.
        /// </summary>
        public bool Override
        {
            get
            {
                return ((_coverage >> 3) & 0x1) == 1;
            }
        }

        /// <summary>
        /// Format of the subtable. Only formats 0 and 2 have been defined. 
        /// Formats 1 and 3 through 255 are reserved for future use.
        /// </summary>
        public ushort Format
        {
            get
            {
                var b = (byte)((_coverage >> 7) & 0xff);
                return BitConverter.ToUInt16(new byte[] { b, byte.MinValue });
            }
        }
    }
}
