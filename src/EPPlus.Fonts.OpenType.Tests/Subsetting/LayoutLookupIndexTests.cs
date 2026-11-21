using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.IO;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class LayoutLookupIndexTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        [DataRow("Mulish-Regular.ttf")]
        [DataRow("Roboto-Regular.ttf")]
        [DataRow("OpenSans-Regular.ttf")]
        [DataRow("EBGaramond-Regular.ttf")]
        [DataRow("CrimsonText-Regular.ttf")]
        public void Subset_GsubLookupIndices_AreWithinLookupList(string fileName)
        {
            var font = new OpenTypeFont(File.ReadAllBytes(Path.Combine(FontFolder, fileName)));

            // Control: the original font must pass, otherwise the loader is at fault, not the subsetter
            FontTestHelper.AssertGsubLookupIndicesValid(font);

            var subsetFont = font.CreateSubset("Hello World 123");

            // Serialize and reload so the assertion checks what OTS and browsers will read
            var bytes = new OpenTypeFontSerializer(subsetFont).Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            SaveFontForCurrentTest(parsedFont, Path.GetFileNameWithoutExtension(fileName));

            FontTestHelper.AssertGsubLookupIndicesValid(parsedFont);
        }

        [TestMethod]
        public void Subset_GsubExtensionFont_MatchesPlainFont()
        {
            var plainFont = new OpenTypeFont(File.ReadAllBytes(Path.Combine(FontFolder, "Mulish-Regular.ttf")));
            var extensionFont = new OpenTypeFont(File.ReadAllBytes(Path.Combine(FontFolder, "Mulish-Regular-GsubExtension.ttf")));

            // Precondition: the test font must actually use Extension lookups
            Assert.IsTrue(FontTestHelper.HasGsubLookupType(extensionFont, 7), "Test font should contain GSUB Extension lookups");

            const string text = "Hello World 123";
            var plainSubset = new OpenTypeFont(new OpenTypeFontSerializer(plainFont.CreateSubset(text)).Serialize());
            var extensionSubset = new OpenTypeFont(new OpenTypeFontSerializer(extensionFont.CreateSubset(text)).Serialize());

            SaveFontForCurrentTest(extensionSubset);

            Assert.IsNotNull(extensionSubset.GsubTable, "GSUB should survive subsetting of a font with Extension lookups");

            // Discovery through Extension must find the same glyphs
            Assert.AreEqual(plainSubset.MaxpTable.numGlyphs, extensionSubset.MaxpTable.numGlyphs, "Glyph count differs");

            // Extension lookups are written with their inner type, so the result must be identical
            var expected = DescribeGsubLookups(plainSubset);
            var actual = DescribeGsubLookups(extensionSubset);

            Console.WriteLine("Plain subset:");
            expected.ForEach(Console.WriteLine);
            Console.WriteLine("Extension subset:");
            actual.ForEach(Console.WriteLine);

            CollectionAssert.AreEqual(expected, actual, "GSUB of the Extension subset differs from the plain subset");

            FontTestHelper.AssertGsubLookupIndicesValid(extensionSubset);
        }

        /// <summary>
        /// Describes each GSUB lookup by type, subtable count and, for chaining contextual
        /// subtables, their lookup records, so two fonts' lookup lists can be compared.
        /// </summary>
        private static List<string> DescribeGsubLookups(OpenTypeFont font)
        {
            var result = new List<string>();
            foreach (var lookup in font.GsubTable.LookupList.Lookups)
            {
                var sb = new StringBuilder();
                sb.Append("type ").Append(lookup.LookupType).Append(", subtables ").Append(lookup.SubTables.Count);
                foreach (var subTable in lookup.SubTables)
                {
                    if (subTable is ChainingContextualSubstFormat3 chain)
                    {
                        sb.Append(" -> [")
                          .Append(string.Join(", ", chain.SubstLookupRecords.Select(r => r.SequenceIndex + ":" + r.LookupListIndex).ToArray()))
                          .Append("]");
                    }
                }
                result.Add(sb.ToString());
            }
            return result;
        }
    }
}