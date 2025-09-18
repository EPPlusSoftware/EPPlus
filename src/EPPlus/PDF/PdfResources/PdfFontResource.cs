using FontLab1;
using FontLab1.GenericMeasurements;
using OfficeOpenXml.PDF.PdfObjects.PdfFonts;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfFontResource : PdfResource
    {
        internal string fontName;
        internal int fontObjectNumber = -1;
        internal int descObjectNumber = -1;
        internal int widthObjectNumber = -1;
        internal TtfFont fontData;
        private int firstChar = 32;
        private int lastChar = 255;

        public PdfFontResource(string fontName, string subFamily, int labelNumber, PdfPageSettings pageSettings)
            : base("F", labelNumber)
        {
            this.fontName = fontName;
            fontData = GenericFonts.GetFontData(pageSettings, fontName, subFamily);
        }

        //Get the Font Descriptor object to write in PDF.
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

        //Get the Widths object to write in PDF.
        internal PdfFontWidths GetWidthsObject(int objectNumber, int version = 0)
        {
            List<int> widths = new List<int>();
            int fallbackWidth = fontData.Os2Table.xAvgCharWidth;
            for (int c = firstChar; c <= lastChar; c++)
            {
                fontData.CmapTable.EncodingRecords[0].CharMappingsToGlyphIndex.TryGetValue((char)c, out ushort gi);
                if (gi == 0 && c != 0)
                {
                    int normalizedWidth = (int)System.Math.Round(fallbackWidth / (double)fontData.HeadTable.UnitsPerEm * 1000);
                    widths.Add(normalizedWidth);
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi];
                    var advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                    int normalizedWidth = (int)System.Math.Round(advanceWidth / (double)fontData.HeadTable.UnitsPerEm * 1000);
                    widths.Add(normalizedWidth);
                }
            }
            widthObjectNumber = objectNumber;
            return new PdfFontWidths(objectNumber, widths, version);
        }

        //Get the Font Object to write in PDF.
        internal PdfFont GetFontObject(int objectNumber, int version = 0)
        {
            fontObjectNumber = objectNumber;
            return new PdfFont(objectNumber, fontName, PdfFontSubType.Type1, firstChar, lastChar, widthObjectNumber, descObjectNumber, PdfFontEncoding.WinAnsiEncoding);
        }
    }
}
