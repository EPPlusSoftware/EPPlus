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
