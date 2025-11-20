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
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("hmtx");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.HmtxTable.Serialize(writer);
            var hmtxBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, hmtxBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hmtxBytes);
        }

        [TestMethod]
        public void SerializeHeadTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("head");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.HeadTable.Serialize(writer);
            var headBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, headBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, headBytes);
        }

        [TestMethod]
        public void SerializeMaxpTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("maxp");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.MaxpTable.Serialize(writer);
            var maxpBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, maxpBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, maxpBytes);
        }

        [TestMethod]
        public void SerializeHheaTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("hhea");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.HheaTable.Serialize(writer);
            var hheaBytes = ms.ToArray();

           Assert.AreEqual(originalBytes.Length, hheaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hheaBytes);
        }

        [TestMethod]
        public void SerializePostTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("post");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.PostTable.Serialize(writer);
            var postBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, postBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, postBytes);
        }

        [TestMethod]
        public void SerializeNameTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("name");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.NameTable.Serialize(writer);
            var nameBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, nameBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, nameBytes);
        }

        [TestMethod]
        [DataRow("Roboto", "Regular", 96)]
        //[DataRow("Roboto", "Italic", 96)]
        //[DataRow("EB Garamond", "Regular", 100)]
        //[DataRow("Mulish", "Regular", 100)]
        public void SerializeOs2Table(string fontName, string subFamily, int expectedLength)
        {
            var sf = FontScanner.ScanFor(_fontFolder, fontName, subFamily);
            var originalBytes = sf.GetTableBytes("OS/2");

            var font = OpenTypeFonts.GetFontData(_fontFolders, fontName, subFamily, false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.Os2Table.Serialize(writer);
            var os2Bytes = ms.ToArray();

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
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("loca");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font.LocaTable.Serialize(writer);
            var locaBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, locaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, locaBytes);
        }

        [TestMethod]
        public void SerializeCmapTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("cmap");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font.CmapTable.Serialize(writer);
            var cmapBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, cmapBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, cmapBytes);
        }

        [TestMethod]
        public void SerializeCmapTable_Format12()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Noto Emoji", "Regular");
            var originalBytes = sf.GetTableBytes("cmap");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Noto Emoji", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font.CmapTable.Serialize(writer);
            var cmapBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, cmapBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, cmapBytes);
        }

        [TestMethod]
        [DataRow("Roboto", "Regular")]
        [DataRow("Noto Emoji", "Regular")]
        [DataRow("EB Garamond", "Regular")]
        [DataRow("Mulish", "Regular")]
        public void SerializeGlyfTable(string fontName, string subFamily, string fontFolder = "")
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
            
            var sf = FontScanner.ScanFor(fontFolder, fontName, subFamily);
            var originalBytes = sf.GetTableBytes("glyf");

            var font = OpenTypeFonts.GetFontData(fontFolders, fontName, subFamily, false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font.GlyfTable.Serialize(writer);
            var glyfBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, glyfBytes?.Length, $"Font {fontName} {subFamily} Length differ ");
            CollectionAssert.AreEqual(originalBytes, glyfBytes, $"Font {fontName} {subFamily} Bytes differ ");
        } 

        [TestMethod]
        [Ignore("Was only able to find fonts with kern table among Windows fonts. These cannot be distributed with the test project due to licensing.")]
        public void SerializeKernTable()
        {
            var sf = FontScanner.ScanFor(@"c:\windows\fonts", "Arial", "Regular");
            var originalBytes = sf.GetTableBytes("kern");

            var font = OpenTypeFonts.GetFontData(new List<string> { @"c:\windows\fonts" } , "Arial", "Regular", false);
            var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);
            font?.KernTable.Serialize(writer);
            var kernBytes = ms.ToArray();

            Assert.AreEqual(originalBytes.Length, kernBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, kernBytes);
        }

        [TestMethod]
        public void FindKernTable()
        {
            var fonts = FontScanner.GetAllScannedFontsInPath(@"c:\windows\fonts");
            var kernFonts = new List<ScannedFont>();
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
