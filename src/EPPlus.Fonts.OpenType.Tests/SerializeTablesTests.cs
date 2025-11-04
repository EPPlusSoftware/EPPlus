using EPPlus.Fonts.OpenType.Scanner;
using System;
using System.Collections.Generic;
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
            var hmtxBytes = font?.HmtxTable.Serialize();

            Assert.AreEqual(originalBytes.Length, hmtxBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hmtxBytes);
        }

        [TestMethod]
        public void SerializeHeadTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("head");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var headBytes = font?.HeadTable.Serialize();

            Assert.AreEqual(originalBytes.Length, headBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, headBytes);
        }

        [TestMethod]
        public void SerializeMaxpTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("maxp");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var maxpBytes = font?.MaxpTable.Serialize();

            Assert.AreEqual(originalBytes.Length, maxpBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, maxpBytes);
        }

        [TestMethod]
        public void SerializeHheaTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("hhea");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var hheaBytes = font?.HheaTable.Serialize();

            Assert.AreEqual(originalBytes.Length, hheaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hheaBytes);
        }

        [TestMethod]
        public void SerializePostTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("post");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var postBytes = font?.PostTable.Serialize();

            Assert.AreEqual(originalBytes.Length, postBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, postBytes);
        }

        [TestMethod]
        public void SerializeNameTable()
        {
            var sf = FontScanner.ScanFor(_fontFolder, "Roboto", "Regular");
            var originalBytes = sf.GetTableBytes("name");

            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var nameBytes = font?.NameTable.Serialize();

            var failIndex = -1;
            for (var i = 0; i < originalBytes.Length; i++)
            {
                var ob = originalBytes[i];
                var nb = nameBytes[i];
                if(ob != nb)
                {
                    failIndex = i;
                    break;
                }
            }
            var obAtFail = originalBytes[failIndex];
            var nbAtFail = nameBytes[failIndex];
            Assert.AreEqual(-1, failIndex);
            //Assert.AreEqual(originalBytes.Length, nameBytes?.Length);
            //CollectionAssert.AreEqual(originalBytes, nameBytes);
        }

        [TestMethod]
        [Ignore("Was only able to find fonts with kern table among Windows fonts. These cannot be distributed with the test project due to licensing.")]
        public void SerializeKernTable()
        {
            var sf = FontScanner.ScanFor(@"c:\windows\fonts", "Arial", "Regular");
            var originalBytes = sf.GetTableBytes("kern");

            var font = OpenTypeFonts.GetFontData(new List<string> { @"c:\windows\fonts" } , "Arial", "Regular", false);
            var kernBytes = font?.KernTable.Serialize();

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
