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
using EPPlus.Export.Pdf.DocumentObjects.Fonts;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Export.Pdf.Resources
{
    internal class PdfFontResource : PdfResource
    {
        internal string fontName;
        internal int fontObjectNumber = -1;
        internal int CIDFontObjectNumber = -1;
        internal int type0FontObjectNumber = -1;
        internal int unicodeCMapFontObjectNumber = -1;
        internal int embedFontStreamObjectNumber = -1;
        internal int fontDescObjectNumber = -1;
        internal int fontWidthObjectNumber = -1;
        internal int cidSetObjectNumber = -1;
        internal OpenTypeFont fontData;
        private OpenTypeFontEngine _fontEngine;
        private int firstChar = 32;
        private int lastChar = 255;
        private CIDSystemInfo cidSystemInfo = null;
        internal string type0Encoding = "Identity-H";
        internal TextShaper Shaper = null;
        internal ShapedText ShapedText;
        internal HashSet<char> Subset = new HashSet<char>();
        internal HashSet<ushort> Gids = new HashSet<ushort>();
        internal Dictionary<ushort, string> charactermappings = new Dictionary<ushort, string>();
        internal FontSubsetManager fontSubsetManager;

        public PdfFontResource(string fontName, FontSubFamily subFamily, int labelNumber, PdfPageSettings pageSettings)
            : base("F", labelNumber)
        {
            this.fontName = fontName;
            _fontEngine = pageSettings.FontEngine;
            fontData = _fontEngine.LoadFont(fontName, subFamily);
            fontSubsetManager = new FontSubsetManager(pageSettings.FontEngine, fontData);
        }

        //Get the Font Descriptor object to write in PDF.
        internal PdfFontDescriptor GetFontDescriptorObject(int objectNumber, int version = 0)
        {
            int flag = 0;
            if (fontData.PostTable.isFixedPitch != 0)
                flag |= 1;
            if (fontData.GetEnglishFontFamilyName().ToLower().Contains("serif"))
                flag |= 1 << 1;
            bool isSymbolic = false;
            foreach (var sub in fontData.CmapTable.EncodingRecords)
            {
                if (sub.PlatformId == Fonts.OpenType.Tables.Cmap.Platforms.Windows)
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
                if (sub.PlatformId == Fonts.OpenType.Tables.Cmap.Platforms.Macintosh || sub.PlatformId == Fonts.OpenType.Tables.Cmap.Platforms.Unicode)
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
            if (fontData.PostTable.italicAngle.RawValue != 0 || (fontData.Os2Table.fsSelection & FsSelectionFlags.Italic) != 0)
                flag |= 1 << 6;
            if (((ushort)fontData.Os2Table.fsSelection & 0x100) != 0)
                flag |= 1 << 16;
            if (((ushort)fontData.Os2Table.fsSelection & 0x200) != 0)
                flag |= 1 << 17;
            if (((ushort)fontData.Os2Table.fsSelection & 0x400) != 0)
                flag |= 1 << 18;
            var fontBBox = new Rect();
            fontBBox.X = fontData.HeadTable.Xmin;
            fontBBox.Y = fontData.HeadTable.Ymin;
            fontBBox.Width = fontData.HeadTable.Xmax;
            fontBBox.Height = fontData.HeadTable.Ymax;
            fontDescObjectNumber = objectNumber;
            return new PdfFontDescriptor
            (
                objectNumber,
                fontData.FullName,
                flag,
                fontBBox,
                Convert.ToDouble(fontData.PostTable.italicAngle.FloatValue),
                fontData.Os2Table.sTypoAscender,
                fontData.Os2Table.sTypoDescender,
                0,
                fontData.Os2Table.sCapHeight,
                embedFontStreamObjectNumber,
                cidSetObjectNumber,
                version
            );
        }

        //Get the Widths object to write in PDF.
        internal PdfFontWidths GetWidthsObject(int objectNumber, int version = 0)
        {
            List<int> widths = new List<int>();
            int fallbackWidth = fontData.Os2Table.xAvgCharWidth;
            var glyphMappings = fontData.CmapTable.GetPreferredSubtable().GetGlyphMappings();
            for (int c = firstChar; c <= lastChar; c++)
            {
                var gi = glyphMappings.GetGlyphIndex((char)c);
                if (gi == 0 && c != 0)
                {
                    int normalizedWidth = (int)System.Math.Round(fallbackWidth / (double)fontData.HeadTable.UnitsPerEm * 1000);
                    widths.Add(normalizedWidth);
                }
                else
                {
                    var hhMetric = fontData.HmtxTable.hMetrics[gi ?? 0];
                    var advanceWidth = Convert.ToInt16(hhMetric.advanceWidth);
                    int normalizedWidth = (int)System.Math.Round(advanceWidth / (double)fontData.HeadTable.UnitsPerEm * 1000);
                    widths.Add(normalizedWidth);
                }
            }
            fontWidthObjectNumber = objectNumber;
            return new PdfFontWidths(objectNumber, widths, version);
        }

        //Get the Font Object to write in PDF.
        internal PdfFont GetFontObject(int objectNumber, int version = 0)
        {
            fontObjectNumber = objectNumber;
            return new PdfFont(objectNumber, fontData.FullName, PdfFontSubType.Type1, firstChar, lastChar, fontWidthObjectNumber, fontDescObjectNumber, PdfFontEncoding.WinAnsiEncoding);
        }

        internal PdfCIDFont GetCIDFontObject(int objectNumber, int version = 0)
        {
            CIDFontObjectNumber = objectNumber;
            if (cidSystemInfo == null)
            {
                cidSystemInfo = new CIDSystemInfo();
            }
            cidSystemInfo.Registry = "Adobe";
            cidSystemInfo.Ordering = "Identity";
            cidSystemInfo.Supplement = 0;
            return new PdfCIDFont(objectNumber, fontData, Gids, CIDFontSubtype.CIDFontType2, cidSystemInfo, "Identity", fontDescObjectNumber);
        }

        internal PdfType0FontDict GetType0FontDictObject(int objectNumber, int version = 0)
        {
            type0FontObjectNumber = objectNumber;
            return new PdfType0FontDict(objectNumber, fontData.FullName, type0Encoding, CIDFontObjectNumber, unicodeCMapFontObjectNumber);
        }

        internal PdfToUnicodeCMap GetUnicodeCmapObject(int objectNumber, int version = 0)
        {
            unicodeCMapFontObjectNumber = objectNumber;
            return new PdfToUnicodeCMap(objectNumber, charactermappings);
        }

        internal PdfFontStream GetEmbeddedFontStreamObject(int objectNumber, int version = 0)
        {
            embedFontStreamObjectNumber = objectNumber;
            return new PdfFontStream(objectNumber, fontData, version);
        }

        internal PdfCidSet GetCidSet(int objectNumber, int version = 0)
        {
            if (!fontData.IsSubset || Gids.Count == 0)
                return null;
            int maxGid = Gids.Max();
            int numBytes = (maxGid / 8) + 1;
            var cidSet = new byte[numBytes];
            foreach (var gid in Gids)
            {
                int byteIndex = gid / 8;
                int bitIndex = 7 - (gid % 8);
                cidSet[byteIndex] |= (byte)(1 << bitIndex);
            }
            cidSetObjectNumber = objectNumber;
            return new PdfCidSet(objectNumber, cidSet, version);
        }
    }
}
