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
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts
{
    [DebuggerDisplay("x: ({Xmin} to {Xmax}), y: ({Ymin} to {Ymax})")]
    internal struct BoundingRectangle
    {
        public BoundingRectangle(short xMin, short yMin, short xMax, short yMax)
        {
            Xmin = xMin;
            Ymin = yMin;
            Xmax = xMax;
            Ymax = yMax;
        }
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
    }
}
