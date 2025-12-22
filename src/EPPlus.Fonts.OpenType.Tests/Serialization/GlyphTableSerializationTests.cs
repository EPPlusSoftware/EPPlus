using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    [TestClass]
    public class GlyphTableSerializationTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void SerializeLocaTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("loca");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var locaBytes = font.LocaTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, locaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, locaBytes);
        }

        [TestMethod]
        public void SerializeHtmxTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("hmtx");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var hmtxBytes = font?.HmtxTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, hmtxBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hmtxBytes);
        }

        [TestMethod]
        [DataRow("Roboto", FontSubFamily.Regular)]
        [DataRow("Noto Emoji", FontSubFamily.Regular)]
        [DataRow("EB Garamond", FontSubFamily.Regular)]
        [DataRow("Mulish", FontSubFamily.Regular)]
        public void SerializeGlyfTable(string fontName, FontSubFamily subFamily)
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolders, fontName, subFamily);
            var originalBytes = ffi.GetTableBytes("glyf");

            var font = OpenTypeFonts.GetFontData(FontFolders, fontName, subFamily, false, true);
            var glyfBytes = font.GlyfTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, glyfBytes?.Length,
                $"Font {fontName} {subFamily} Length differ");
            CollectionAssert.AreEqual(originalBytes, glyfBytes,
                $"Font {fontName} {subFamily} Bytes differ");
        }
    }
}
