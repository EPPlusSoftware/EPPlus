using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.PdfPageSettings;

namespace OfficeOpenXml.PDF.PdfFontData
{
    internal class PdfFontProperties
    {
        public PdfFontProperties()
        {
            ClassWidths = new Dictionary<FontMetricsClass, float>();
            CharMetrics = new Dictionary<char, FontMetricsClass>();
        }

        public FontMetricsFamilies Family { get; set; }

        public FontSubFamilies SubFamily { get; set; }

        public ushort Version { get; set; }
        public uint FontKey { get; set; }
        public float LineHeight1em { get; set; }

        public FontMetricsClass DefaultWidthClass { get; set; }


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

        public int flags;
        public PdfRect fontBBox;
        public double italicAngle;
        public int ascent;
        public int descent;
        public double stemV;
        public short capheight;
        public int firstChar;
        public int lastChar;

        public static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }
    }
}
