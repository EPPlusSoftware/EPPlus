using FontLab1;
using FontLab1.GenericMeasurements;
using FontLab1.Tables.Cmap;
using FontLab1.Tables.Kern;
using Microsoft.VisualBasic;
using OfficeOpenXml.PDF.PdfObjects;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace OfficeOpenXml.PDF.PdfFontData
{
    internal class PdfFontResource
    {
        internal string fontName;
        internal string labelPrefix = "F";
        internal int labelNumber;
        internal string Label
        { 
            get 
            { 
                return labelPrefix + labelNumber; 
            }
        }
        internal int fontObjectNumber = -1;
        internal int descObjectNumber = -1;
        internal int widthObjectNumber = -1;
        internal TtfFont fontData;

        private int firstChar = 32;
        private int lastChar = 255;

        public PdfFontResource(string fontName, string subFamily, int label, PdfPageSettings pageSettings)
        {
            this.fontName = fontName;
            this.labelNumber = label;
            fontData = GenericFonts.GetFontData(pageSettings, fontName, subFamily);
        }

        internal PdfFontDescriptor GetFontDescriptorObject(int objectNumber, int version = 0)
        {
            int flag = 0;
            if (fontData.postTable.isFixedPitch != 0)
                flag |= 1;
            if (fontData.GetEnglishFontFamilyName().ToLower().Contains("serif"))
                flag |= 1 << 1;
            bool isSymbolic = false;
            foreach (var sub in fontData.CmapTable.EncodingRecords)
            {
                if (sub.PlatformId == FontLab1.Tables.Cmap.Platforms.Windows)
                {
                    if (sub.EncodingId == 0)
                    {
                        isSymbolic = true;
                        break;
                    }
                    if (sub.EncodingId == 1 || sub.EncodingId == 10)
                    {
                        isSymbolic = false;
                        break;
                    }
                }
                if (sub.PlatformId == FontLab1.Tables.Cmap.Platforms.Macintosh || sub.PlatformId == FontLab1.Tables.Cmap.Platforms.Unicode)
                {
                    isSymbolic = true;
                    break;
                }
            }
            if (isSymbolic)
                flag |= 1 << 2; // Symbolic
            else
                flag |= 1 << 5; // Nonsymbolic
            if (fontData.GetEnglishFontFamilyName().ToLower().Contains("script") || fontData.GetEnglishFontFamilyName().ToLower().Contains("cursive"))
                flag |= 1 << 3;
            if (fontData.postTable.italicAngle != 0 || (fontData.Os2Table.fsSelection & 0x01) != 0)
                flag |= 1 << 6;
            if ((fontData.Os2Table.fsSelection & 0x100) != 0)
                flag |= 1 << 16;
            if ((fontData.Os2Table.fsSelection & 0x200) != 0)
                flag |= 1 << 17;
            if ((fontData.Os2Table.fsSelection & 0x400) != 0)
                flag |= 1 << 18;
            var fontBBox = new PdfRect();
            fontBBox.X = fontData.HeadTable.Xmin;
            fontBBox.Y = fontData.HeadTable.Ymin;
            fontBBox.Width = fontData.HeadTable.Xmax;
            fontBBox.Height = fontData.HeadTable.Ymax;
            descObjectNumber = objectNumber;
            return new PdfFontDescriptor
            (
                objectNumber, 
                fontName,
                flag,
                fontBBox,
                fontData.postTable.italicAngle,
                fontData.Os2Table.sTypoAscender,
                fontData.Os2Table.sTypoDescender,
                0,
                fontData.Os2Table.sCapHeight,
                version
            );
        }

        internal PdfFontWidths GetWidthsObject(int objectNumber, int version = 0)
        {
            List<int> widths = new List<int>();
            int fallbackWidth = fontData.Os2Table.xAvgCharWidth;
            for (int c = firstChar; c <= lastChar; c++)
            {
                fontData.CmapTable.EncodingRecords[0].CharMappingsToGlyphIndex.TryGetValue((char)c, out ushort gi);
                if (gi == 0 && c != 0)
                {
                    int normalizedWidth = (int)System.Math.Round((fallbackWidth / (double)fontData.HeadTable.UnitsPerEm) * 1000);
                    widths.Add(normalizedWidth);
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi];
                    var advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                    int normalizedWidth = (int)System.Math.Round((advanceWidth / (double)fontData.HeadTable.UnitsPerEm) * 1000);
                    widths.Add(normalizedWidth);
                }
            }
            widthObjectNumber = objectNumber;
            return new PdfFontWidths(objectNumber, widths, version);
        }

        internal PdfFont GetFontObject(int objectNumber, int version = 0)
        {
            fontObjectNumber = objectNumber;
            return new PdfFont(objectNumber, fontName, PdfFontSubType.Type1, firstChar, lastChar, widthObjectNumber, descObjectNumber, PdfFontEncoding.WinAnsiEncoding);
        }

        internal double MeasureText(string text, double fontSize)
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
                    int kerning = GetKerningAdjustment(lastGlyphIndex, gi);
                    totalAdvanceWidth += kerning;
                }

                lastGlyphIndex = gi;
                firstChar = false;
            }

            // Convert to points
            return (totalAdvanceWidth / (double)fontData.HeadTable.UnitsPerEm) * fontSize;
        }

        int GetKerningAdjustment(ushort left, ushort right)
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
