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
    }
}
