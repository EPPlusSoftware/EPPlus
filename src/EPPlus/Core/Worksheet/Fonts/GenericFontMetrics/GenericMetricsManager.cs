using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal static class GenericMetricsManager
    {
        public static GenericFontMetrics CreateMetrics(SerializedFontMetrics font)
        {
            var min = font.AdvanceWidths.Values.Where(x => x > 0).Min();
            var max = font.AdvanceWidths.Values.Max();
            var level0 = min;
            var level1 = (max - min) * 0.15;
            var level2 = (max - min) * 0.3;
            var level3 = (max - min) * 0.45;
            var level4 = (max - min) * 0.6;
            var level5 = (max - min) * 0.75;
            var level6 = (max - min) * 0.875;
            var l1 = new HashSet<char>();
            var l2 = new HashSet<char>();
            var l3 = new HashSet<char>();
            var l4 = new HashSet<char>();
            var l5 = new HashSet<char>();
            var l6 = new HashSet<char>();
            var l7 = new HashSet<char>();
            foreach (var c in font.AdvanceWidths.Keys)
            {
                var width = font.AdvanceWidths[c]; 
                if(width < level1)
                {
                    l1.Add(c);
                }
                else if(width < level2)
                {
                    l2.Add(c);
                }
                else if(width < level3)
                {
                    l3.Add(c);
                }
                else if(width < level4)
                {
                    l4.Add(c);
                }
                else if(width < level5)
                {
                    l5.Add(c);
                }
                else if(width < level6)
                {
                    l6.Add(c);
                }
                else
                {
                    l7.Add(c);
                }
            }

            // set widths per level
            short l1w = Convert.ToInt16((level1 - level0) / 2);
            var l2w = Convert.ToInt16(level1 + (level2 - level1) / 2);
            var l3w = Convert.ToInt16(level2 + (level3 - level2) / 2);
            var l4w = Convert.ToInt16(level3 + (level4 - level3) / 2);
            var l5w = Convert.ToInt16(level4 + (max - level4) / 2);
            var l6w = Convert.ToInt16(level5 + (max - level5) / 2);
            var l7w = Convert.ToInt16(level6 + (max - level6) / 2);

            var metrics = new GenericFontMetrics();
            metrics.ClassWidths.Add(FontMetricsClass.Class1, l1w);
            metrics.ClassWidths.Add(FontMetricsClass.Class2, l2w);
            metrics.ClassWidths.Add(FontMetricsClass.Class3, l3w);
            metrics.ClassWidths.Add(FontMetricsClass.Class4, l4w);
            metrics.ClassWidths.Add(FontMetricsClass.Class5, l5w);
            metrics.ClassWidths.Add(FontMetricsClass.Class6, l6w);
            metrics.ClassWidths.Add(FontMetricsClass.Class7, l7w);

            foreach (var c1 in l1)
            {
                metrics.CharMetrics.Add(c1, FontMetricsClass.Class1);
            }
            foreach (var c2 in l2)
            {
                metrics.CharMetrics.Add(c2, FontMetricsClass.Class2);
            }
            foreach (var c3 in l3)
            {
                metrics.CharMetrics.Add(c3, FontMetricsClass.Class3);
            }
            foreach (var c4 in l4)
            {
                metrics.CharMetrics.Add(c4, FontMetricsClass.Class4);
            }
            foreach (var c5 in l5)
            {
                metrics.CharMetrics.Add(c5, FontMetricsClass.Class5);
            }
            foreach (var c6 in l6)
            {
                metrics.CharMetrics.Add(c6, FontMetricsClass.Class6);
            }
            foreach (var c7 in l7)
            {
                metrics.CharMetrics.Add(c7, FontMetricsClass.Class7);
            }
            metrics.LineHeight = font.LineHeight;
            return metrics;
        }
    }
}
