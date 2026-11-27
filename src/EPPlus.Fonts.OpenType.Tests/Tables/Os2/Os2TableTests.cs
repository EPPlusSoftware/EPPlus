using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tests.Tables
{
    [TestClass]
    public class Os2TableTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        [DataRow("Mulish-Regular.ttf")]
        [DataRow("Roboto-Regular.ttf")]
        [DataRow("OpenSans-Regular.ttf")]
        [DataRow("EBGaramond-Regular.ttf")]
        [DataRow("CrimsonText-Regular.ttf")]
        public void Os2_Subset_HasLengthMatchingVersion(string fileName)
        {
            var font = new OpenTypeFont(File.ReadAllBytes(Path.Combine(FontFolder, fileName)));
            ushort version = font.Os2Table.version;
            Console.WriteLine(fileName + ": OS/2 version " + version);

            if (version < 5)
            {
                // Optical size fields only exist from version 5; anything else was read past the table end
                Assert.AreEqual((ushort)0, font.Os2Table.usLowerOpticalPointSize, "usLowerOpticalPointSize read for version " + version);
                Assert.AreEqual((ushort)0, font.Os2Table.usUpperOpticalPointSize, "usUpperOpticalPointSize read for version " + version);
            }

            var bytes = new OpenTypeFontSerializer(font.CreateSubset("Hello World 123")).Serialize();
            SaveFontForCurrentTest(new OpenTypeFont(bytes), Path.GetFileNameWithoutExtension(fileName));

            Assert.AreEqual(GetExpectedOs2Length(version), GetTableLength(bytes, "OS/2"), "OS/2 length for version " + version);
        }

        private static int GetExpectedOs2Length(ushort version)
        {
            if (version == 0) return 78;
            if (version == 1) return 86;
            if (version <= 4) return 96;
            return 100;
        }

        /// <summary>
        /// Reads a table's length directly from the serialized table directory, so the test
        /// does not depend on the font model.
        /// </summary>
        private static int GetTableLength(byte[] fontBytes, string tag)
        {
            int numTables = (fontBytes[4] << 8) | fontBytes[5];
            for (int i = 0; i < numTables; i++)
            {
                int recordStart = 12 + i * 16;
                string recordTag = System.Text.Encoding.ASCII.GetString(fontBytes, recordStart, 4);
                if (recordTag == tag)
                {
                    return (fontBytes[recordStart + 12] << 24) | (fontBytes[recordStart + 13] << 16)
                         | (fontBytes[recordStart + 14] << 8) | fontBytes[recordStart + 15];
                }
            }
            Assert.Fail("Table " + tag + " not found");
            return 0;
        }
    }
}