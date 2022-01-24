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
using OfficeOpenXml.Core.Worksheet.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Head
{
    /// <summary>
    /// This table gives global information about the font.
    /// </summary>
    internal class HeadTable
    {
        public enum IndexToLocFormats : short
        {
            Offset16 = 0,
            Offset32 = 1
        }
        public ushort MajorVersion { get; set; }

        public ushort MinorVersion { get; set; }

        /// <summary>
        /// Set to a value from 16 to 16384. Any value in this range is valid. In fonts that have TrueType outlines, a power of 2 is recommended as this allows performance optimizations in some rasterizers.
        /// </summary>
        public ushort UnitsPerEm { get; set; }

        /// <summary>
        /// Minimum x coordinate across all glyph bounding boxes.
        /// </summary>
        public short Xmin { get; set; }

        /// <summary>
        /// Minimum y coordinate across all glyph bounding boxes.
        /// </summary>
        public short Ymin { get; set; }

        /// <summary>
        /// Maximum x coordinate across all glyph bounding boxes.
        /// </summary>
        public short Xmax { get; set; }

        /// <summary>
        /// Maximum y coordinate across all glyph bounding boxes.
        /// </summary>
        public short Ymax { get; set; }

        /// <summary>
        /// Smallest readable size in pixels.
        /// </summary>
        public ushort LowestRecPPEM { get; set; }

        /// <summary>
        /// 0 for short offsets (Offset16), 1 for long (Offset32).
        /// </summary>
        public IndexToLocFormats IndexToLocFormat { get; set; }

        public BoundingRectangle GetDefaultBounds()
        {
            return new BoundingRectangle(Xmin, Ymin, Xmax, Ymax);
        }
    }
}
