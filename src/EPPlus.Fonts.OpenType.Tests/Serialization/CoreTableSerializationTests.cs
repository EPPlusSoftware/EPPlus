using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    [TestClass]
    public class CoreTableSerializationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void SerializeHeadTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("head");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var headBytes = font?.HeadTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, headBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, headBytes);
        }

        [TestMethod]
        public void SerializeMaxpTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("maxp");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var maxpBytes = font?.MaxpTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, maxpBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, maxpBytes);
        }

        [TestMethod]
        public void SerializeHheaTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("hhea");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var hheaBytes = font?.HheaTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, hheaBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, hheaBytes);
        }

        [TestMethod]
        public void SerializePostTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolders, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("post");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var postBytes = font?.PostTable.Serialize(font);

            Assert.AreEqual(originalBytes.Length, postBytes?.Length);
            CollectionAssert.AreEqual(originalBytes, postBytes);
        }

        [TestMethod]
        public void SerializeNameTable()
        {
            var ffi = FontScannerV2.FindBestMatch(FontFolder, "Roboto", FontSubFamily.Regular);
            var originalBytes = ffi.GetTableBytes("name");

            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
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
            var ffi = FontScannerV2.FindBestMatch(FontFolder, fontName, subFamily);
            var originalBytes = ffi.GetTableBytes("OS/2");

            var font = OpenTypeFonts.GetFontData(FontFolders, fontName, subFamily, false, true);
            var os2Bytes = font?.Os2Table.Serialize(font);

            Assert.AreEqual(expectedLength, os2Bytes?.Length);
            if (expectedLength > originalBytes.Length)
            {
                os2Bytes = os2Bytes?.Take(originalBytes.Length).ToArray();
            }
            CollectionAssert.AreEqual(originalBytes, os2Bytes);
        }
    }
}
