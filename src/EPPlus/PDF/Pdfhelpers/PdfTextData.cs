using FontLab1;
using FontLab1.GenericMeasurements;
using FontLab1.Tables.Cmap;
using FontLab1.Tables.Kern;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Linq;

namespace OfficeOpenXml.PDF.Pdfhelpers
{
    internal static class PdfTextData
    {
        internal static TtfFont GetFontData(PdfPageSettings pageSettings, string fontName, string subFamily)
        {
            return GenericFonts.GetFontData(pageSettings, fontName, subFamily);
        }

        internal static double MeasureFontHeight(TtfFont font, double fontSize)
        {
            var asc = font.Os2Table.usWinAscent;
            var desc = font.Os2Table.usWinDescent;
            //var gap = font.Os2Table.sTypoLineGap;
            var size = fontSize;
            var em = font.HeadTable.UnitsPerEm;
            //var lineHeight = asc + System.Math.Abs(desc) + gap;
            var lineHeight = asc - desc;
            var lineHeightPt = lineHeight * (size / em);
            return lineHeightPt;
        }

        internal static double MeasureAscent(TtfFont font, double fontSize)
        {
            return font.Os2Table.usWinAscent * (fontSize / font.HeadTable.UnitsPerEm);
        }

        internal static double MeasureDescent(TtfFont font, double fontSize)
        {
            return font.Os2Table.usWinDescent * (fontSize / font.HeadTable.UnitsPerEm);
        }

        internal static double MeasureText(string text, double fontSize, TtfFont fontData)
        {
            //double totalAdvanceWidth = 0;

            //foreach (char c in text)
            //{
            //    ushort gi = fontData.CmapTable.EncodingRecords[0].Mappings[c].GlyphIndex;
            //    int advanceWidth;
            //    if (gi == 0 && c != 0)
            //    {
            //        advanceWidth = fontData.Os2Table.xAvgCharWidth;
            //    }
            //    else
            //    {
            //        var hhMetric = fontData.HmtxTable.hMetrics[gi];
            //        advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
            //    }
            //    totalAdvanceWidth += advanceWidth;
            //}
            //return (totalAdvanceWidth / fontData.HeadTable.UnitsPerEm) * fontSize;

            double totalAdvanceWidth = 0;
            ushort lastGlyphIndex = 0;
            bool firstChar = true;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                var encodingRecord = fontData.CmapTable.EncodingRecords.FirstOrDefault(er => er.PlatformId == Platforms.Windows && er.EncodingId == 1);

                if (encodingRecord == null)
                    throw new Exception("Could not find Microsoft Unicode cmap (PlatformID 3, EncodingID 1).");

                GlyphMapping[] mappings = encodingRecord.Mappings;



                encodingRecord.CharMappingsToGlyphIndex.TryGetValue(c, out ushort gi);

                //ushort gi = fontData.CmapTable.EncodingRecords[0].Mappings[c].GlyphIndex;
                int advanceWidth;
                if (gi == 0 && c != 0)
                {
                    advanceWidth = fontData.Os2Table.xAvgCharWidth;
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi];
                    advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                }
                totalAdvanceWidth += advanceWidth;

                // Kerning adjustment
                if (!firstChar)
                {
                    int kerning = GetKerningAdjustment(lastGlyphIndex, gi, fontData);
                    totalAdvanceWidth += kerning;
                }

                lastGlyphIndex = gi;
                firstChar = false;
            }

            // Convert to points
            return (totalAdvanceWidth / (double)fontData.HeadTable.UnitsPerEm) * fontSize;
        }

        private static int GetKerningAdjustment(ushort left, ushort right, TtfFont fontData)
        {
            foreach (var subtable in fontData.KernTable.SubTables)
            {
                if (subtable.Format0Subtable == null)
                    continue;

                // Format 0 only
                int format = subtable.coverage._coverage >> 8;
                bool isHorizontal = (subtable.coverage._coverage & 0x1) == 1;

                if (format != 0 || !isHorizontal)
                    continue;

                KerningPair[] pairs = subtable.Format0Subtable.Pairs;
                if (pairs == null)
                    continue;

                for (int i = 0; i < pairs.Length; i++)
                {
                    if (pairs[i].left == left && pairs[i].right == right)
                        return pairs[i].value;
                }
            }

            return 0;
        }
    }
}
