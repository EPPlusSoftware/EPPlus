/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using System.Collections.Generic;
using EPPlus.Export.Pdf.PdfSettings;

namespace EPPlus.Export.Pdf.PdfFontData
{
    internal class PdfFontProperties
    {
        public FontMetricsFamilies Family { get; set; }

        public FontSubFamilies SubFamily { get; set; }
        public ushort Version { get; set; }

        public uint FontKey { get; set; }

        public float LineHeight1em { get; set; }

        public FontMetricsClass DefaultWidthClass { get; set; }

        public Dictionary<FontMetricsClass, float> ClassWidths { get; private set; }

        public Dictionary<char, FontMetricsClass> CharMetrics { get; private set; }

        public int Flags { get; set; }

        public PdfRect FontBBox { get; set; }

        public double ItalicAngle { get; set; }

        public int Ascent { get; set; }

        public int Descent { get; set; }

        public double StemV { get; set; }

        public short Capheight { get; set; }

        public int FirstChar { get; set; }

        public int LastChar { get; set; }

        public PdfFontProperties()
        {
            ClassWidths = new Dictionary<FontMetricsClass, float>();
            CharMetrics = new Dictionary<char, FontMetricsClass>();
        }

        public static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)(k1 << 16 | k2 & 0xffff);
        }

        public uint GetKey()
        {
            return GetKey(Family, SubFamily);
        }
    }
}
