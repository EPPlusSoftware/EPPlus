using EPPlus.Fonts.OpenType;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Text
{
    internal static class TextWrapper
    {
        internal static List<string> GetLines(string content, FontMeasurerTrueType measurer, double maxWidth = double.NaN)
        {
            if (string.IsNullOrEmpty(content)) return null;

            List<string> lines = new List<string>();

            if (double.IsNaN(maxWidth))
            {
                lines = measurer.MeasureAndWrapText(content, maxWidth);
            }
            else
            {
                lines = content.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();
            }

            return lines;
        }

        internal static List<double> GetContentWidths(string content, FontMeasurerTrueType measurer, double maxWidth = double.NaN)
        {
            var lines = GetLines(content, measurer, maxWidth);
            return GetContentWidths(lines.ToArray(), measurer, maxWidth);
        }

        /// <summary>
        /// Assumes you already have correctly calculated lines
        /// </summary>
        /// <param name="content"></param>
        /// <param name="measurer"></param>
        /// <param name="maxWidth"></param>
        /// <returns></returns>
        internal static List<double> GetContentWidths(string[] content, FontMeasurerTrueType measurer, double maxWidth = double.NaN)
        {
            double[] Widths = new double[content.Length];

            for (int i = 0; i < content.Length; i++)
            {
                Widths[i] = measurer.MeasureTextWidth(content[i]);
            }

            return Widths.ToList();
        }
    }
}
