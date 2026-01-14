using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    [TestClass]
    public class CmapSerializationTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void SerializeCmapTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("cmap");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var cmapBytes = font.CmapTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, cmapBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, cmapBytes);
        }

        [TestMethod]
        public void SerializeCmapTable_Format12()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Noto Emoji", FontSubFamily.Regular, false, true);

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
            var tempFont = new OpenTypeFont(bytes, font.Format);

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
    }
}
