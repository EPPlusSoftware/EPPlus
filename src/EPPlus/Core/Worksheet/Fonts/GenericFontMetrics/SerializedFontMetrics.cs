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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class SerializedFontMetrics
    {
        public SerializedFontMetrics()
        {
            ClassWidths = new Dictionary<FontMetricsClass, float>();
            CharMetrics = new Dictionary<char, FontMetricsClass>();
        }

        public ushort Version { get; set; }
        public FontMetricsFamilies Family { get; set; }
        public FontSubFamilies SubFamily { get; set; }
        public float LineHeight1em { get; set; }

        public float DefaultWidth1em { get; set; }

        public Dictionary<FontMetricsClass, float> ClassWidths
        {
            get;
            private set;
        }

        public Dictionary<char, FontMetricsClass> CharMetrics
        {
            get;
            private set;
        }

        public uint GetKey()
        {
            return GetKey(Family, SubFamily);
        }

        public static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }

    }
}
