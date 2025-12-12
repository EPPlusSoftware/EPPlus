using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class SerializeTablesTests
    {
        private static string _fontFolder = string.Empty;
        private static List<string> _fontFolders = new List<string>();

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
            _fontFolders.Clear();
            _fontFolders.Add(_fontFolder);
            OpenTypeFonts.ClearFontCache();
        }

        [TestMethod]
        public void SerializeHtmxTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("hmtx");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var hmtxBytes = font?.HmtxTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, hmtxBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hmtxBytes);
        }

        [TestMethod]
        public void SerializeHeadTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("head");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var headBytes = font?.HeadTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, headBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, headBytes);
        }

        [TestMethod]
        public void SerializeMaxpTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("maxp");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var maxpBytes = font?.MaxpTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, maxpBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, maxpBytes);
        }

        [TestMethod]
        public void SerializeHheaTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("hhea");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var hheaBytes = font?.HheaTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, hheaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hheaBytes);
        }

        [TestMethod]
        public void SerializePostTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("post");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var postBytes = font?.PostTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, postBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, postBytes);
        }

        [TestMethod]
        public void SerializeNameTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("name");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var nameBytes = font?.NameTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, nameBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, nameBytes);
        }

        [TestMethod]
        [DataRow("Roboto", FontSubFamily.Regular, 96)]
        [DataRow("Roboto", FontSubFamily.Italic, 96)]
        [DataRow("EB Garamond", FontSubFamily.Regular, 100)]
        [DataRow("Mulish", FontSubFamily.Regular, 100)]
        public void SerializeOs2Table(string fontName, FontSubFamily subFamily, int expectedLength)
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, fontName, subFamily);
            var originalBytes = ffi.GetTableBytes("OS/2");

            var font = OpenTypeFonts.GetFontData(_fontFolders, fontName, subFamily, false, true);
            var os2Bytes = font?.Os2Table.Serialize(font);

            Assert.AreEqual(expectedLength, os2Bytes?.Length);
            if(expectedLength > originalBytes.Length)
            {
                os2Bytes = os2Bytes?.Take(originalBytes.Length).ToArray();
            }
            CollectionAssert.AreEqual(originalBytes, os2Bytes);
        }

        [TestMethod]
        public void SerializeLocaTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("loca");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var locaBytes = font.LocaTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, locaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, locaBytes);
        }

        [TestMethod]
        public void SerializeCmapTable()
        {
            var ffi = FontScannerV2.FindBestMatch(_fontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("cmap");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var cmapBytes = font.CmapTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, cmapBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, cmapBytes);
        }

        [TestMethod]
        public void SerializeCmapTable_Format12()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Noto Emoji", FontSubFamily.Regular, false, true);

            // Ta unique chars from the originalfonts cmap (format 4 + 12)
            var allCodePoints = new HashSet<uint>();
            foreach (var sub in font.CmapTable.SubTables)
            {
                if (sub.Format == 14) continue; // skip VS
                var map = sub.GetGlyphMappings().CharCodeToGlyphIndex;
                foreach (var cp in map.Keys)
                    allCodePoints.Add(cp);
            }

            // re-serialize
            var bytes = font.CmapTable.Serialize(font);
            var tempFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Check that ALL original chars are still there
            foreach (uint cp in allCodePoints)
            {
                ushort gid1, gid2;
                bool has1 = font.CmapTable.TryGetGlyphId(cp, out gid1);
                bool has2 = tempFont.CmapTable.TryGetGlyphId(cp, out gid2);

                Assert.IsTrue(has1 == has2, $"Code point U+{cp:X4} lost after roundtrip");
                if (has1)
                    Assert.AreEqual(gid1, gid2);
            }
        }

        [TestMethod]
        [DataRow("Roboto", FontSubFamily.Regular)]
        [DataRow("Noto Emoji", FontSubFamily.Regular)]
        [DataRow("EB Garamond", FontSubFamily.Regular)]
        [DataRow("Mulish", FontSubFamily.Regular)]
        public void SerializeGlyfTable(string fontName, FontSubFamily subFamily, string fontFolder = "")
        {
            var fontFolders = new List<string> { };
            if (string.IsNullOrEmpty(fontFolder))
            {
                fontFolder = _fontFolder;
                fontFolders = new List<string> { fontFolder };
            }
            else
            {
                fontFolders = new List<string> { fontFolder };
            }
            if (string.IsNullOrEmpty(fontFolder)) fontFolder = _fontFolder;
            
            var ffi = FontScannerV2.FindBestMatch(fontFolders, fontName, subFamily);
            var originalBytes = ffi.GetTableBytes("glyf");

            var font = OpenTypeFonts.GetFontData(fontFolders, fontName, subFamily, false, true);
            var glyfBytes = font.GlyfTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, glyfBytes?.Length, $"Font {fontName} {subFamily} Length differ ");
            CollectionAssert.AreEqual(originalBytes, glyfBytes, $"Font {fontName} {subFamily} Bytes differ ");
        } 

        [TestMethod]
        [Ignore("Was only able to find fonts with kern table among Windows fonts. These cannot be distributed with the test project due to licensing.")]
        public void SerializeKernTable()
        {
            var ffi = FontScannerV2.FindBestMatch(@"c:\windows\fonts", "Arial", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("kern");

            var font = OpenTypeFonts.GetFontData(new List<string> { @"c:\windows\fonts" } , "Arial", FontSubFamily.Regular, false);
            var kernBytes = font?.KernTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, kernBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, kernBytes);
        }

        [TestMethod]
        public void FindKernTable()
        {
            var fonts = FontScannerV2.GetAllScannedFontsInPath(@"c:\windows\fonts");
            var kernFonts = new List<FontFaceInfo>();
            foreach (var font in fonts)
            {
                if (font.TableRecords.ContainsKey("kern") || font.TableRecords.ContainsKey("Kern"))
                {
                    kernFonts.Add(font);
                }
            }
            var c = kernFonts.Count;
        }
    }
}
