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
    /// <summary>
    /// This table contains information for horizontal layout. The values in the minRightSidebearing, 
    /// minLeftSideBearing and xMaxExtent should be computed using only glyphs that have contours. 
    /// Glyphs with no contours should be ignored for the purposes of these calculations. All 
    /// reserved areas must be set to 0.
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/hhea
    /// </summary>
    internal class HheaTable
    {
        /// <summary>
        /// Major version number of the horizontal header table — set to 1.
        /// </summary>
        public ushort majorVersion { get; set; }

        /// <summary>
        /// Minor version number of the horizontal header table — set to 0.
        /// </summary>
        public ushort minorVersion { get; set; }

        /// <summary>
        /// Typographic ascent
        /// </summary>
        public short ascender { get; set; }

        /// <summary>
        /// Typographic descent
        /// </summary>
        public short descender { get; set; }

        /// <summary>
        /// Typographic line gap.
        /// Negative LineGap values are treated as zero in some legacy platform implementations.
        /// </summary>
        public short lineGap { get; set; }

        /// <summary>
        /// Maximum advance width value in 'hmtx' table.
        /// </summary>
        public ushort advanceWidthMax { get; set; }

        /// <summary>
        /// Minimum left sidebearing value in 'hmtx' table for glyphs with contours (empty glyphs should be ignored).
        /// </summary>
        public short minLeftSideBearing { get; set; }

        /// <summary>
        /// Minimum right sidebearing value; calculated as min(aw - (lsb + xMax - xMin)) for glyphs with contours (empty glyphs should be ignored).
        /// </summary>
        public short minRightSideBearing { get; set; }

        /// <summary>
        /// Max(lsb + (xMax - xMin)).
        /// </summary>
        public short xMaxExtent { get; set; }

        /// <summary>
        /// Used to calculate the slope of the cursor (rise/run); 1 for vertical.
        /// </summary>
        public short caretSlopeRise { get; set; }

        /// <summary>
        /// 0 for vertical.
        /// </summary>
        public short caretSlopeRun { get; set; }

        /// <summary>
        /// The amount by which a slanted highlight on a glyph needs to be shifted to produce the best appearance. Set to 0 for non-slanted fonts
        /// </summary>
        public short caretOffset { get; set; }

        /// <summary>
        /// 0 for current format.
        /// </summary>
        public short metricDataFormat { get; set; }

        /// <summary>
        /// Number of hMetric entries in 'hmtx' table
        /// </summary>
        public ushort numberOfHMetrics { get; set; }
    }
}
