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
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Kern;
using System;
using System.Linq;

namespace EPPlus.Export.Pdf.Pdfhelpers
{
    internal static class PdfTextData
    {
        internal static OpenTypeFont GetFontData(PdfPageSettings pageSettings, string fontName, string subFamily)
        {
            return OpenTypeFonts.GetFontData(pageSettings.FontDirectories, fontName, subFamily, pageSettings.SearchSystemDirectories);
        }

        internal static double MeasureFontHeight(OpenTypeFont font, double fontSize)
        {
            var asc = font.Os2Table.usWinAscent;
            var desc = font.Os2Table.usWinDescent;
            var size = fontSize;
            var em = font.HeadTable.UnitsPerEm;
            //var lineHeight = asc + System.Math.Abs(desc); //we should use this formula instead, but due to how layouting works we use the other one for now. Need to fix layouting stuff and the issue is when we swtich Y axis most likely.
            var FontHeight = asc - desc;
            var FontHeightPt = FontHeight * (size / em);
            return FontHeightPt;
        }

        internal static double MeasureLineHeight(OpenTypeFont font, double fontSize)
        {
            var asc = font.Os2Table.usWinAscent;
            var desc = font.Os2Table.usWinDescent;
            var lineGap = font.Os2Table.sTypoLineGap;
            var size = fontSize;
            var em = font.HeadTable.UnitsPerEm;
            var lineHeight = asc - desc + lineGap;
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

        internal static double MeasureText(string text, double fontSize, OpenTypeFont fontData)
        {
            double totalAdvanceWidth = 0;
            ushort lastGlyphIndex = 0;
            bool firstChar = true;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                var subTable = fontData.CmapTable.GetSubtable4();
                var gi = subTable.GetGlyphIndex(c);
                //var encodingRecord = fontData.CmapTable.EncodingRecords.FirstOrDefault(er => er.PlatformId == Platforms.Windows && er.EncodingId == 1);
                //if (encodingRecord == null) throw new Exception("Could not find Microsoft Unicode cmap (PlatformID 3, EncodingID 1).");
                //GlyphMapping[] mappings = encodingRecord.Mappings;
                //encodingRecord.CharMappingsToGlyphIndex.TryGetValue(c, out ushort gi);
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
            return totalAdvanceWidth / fontData.HeadTable.UnitsPerEm * fontSize;
        }

        private static int GetKerningAdjustment(ushort left, ushort right, OpenTypeFont fontData)
        {
            foreach (var subtable in fontData.KernTable.SubTables)
            {
                if (subtable.Format0Subtable == null) continue;
                // Format 0 only
                int format = subtable.coverage._coverage >> 8;
                bool isHorizontal = (subtable.coverage._coverage & 0x1) == 1;
                if (format != 0 || !isHorizontal) continue;
                KerningPair[] pairs = subtable.Format0Subtable.Pairs;
                if (pairs == null) continue;
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
