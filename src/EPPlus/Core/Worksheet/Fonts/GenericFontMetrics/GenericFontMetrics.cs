using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class GenericFontMetrics
    {
        public GenericFontMetrics()
        {
            CharMetrics = new Dictionary<char, FontMetricsClass>();
            ClassWidths = new Dictionary<FontMetricsClass, short>();
        }

        public IDictionary<char, FontMetricsClass> CharMetrics
        {
            get; private set;
        }

        public IDictionary<FontMetricsClass, short> ClassWidths
        {
            get; private set;
        }

        public short LineHeight { get; set; }

        public float Measure(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0f;
            var arr = text.ToCharArray();
            var width = 0f;
            foreach(var c in arr)
            {
                var cls = CharMetrics[c];
                var cw = ClassWidths[cls];
                width += cw;
            }
            return width;
        }
    }
}
