using EPPlus.Fonts.OpenType.Tests.Helpers;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    /// <summary>
    /// Generates subsets with demanding typography, for validation with OpenType Sanitizer (OTS).
    /// The saved fonts are checked externally with "python -m ots in.ttf out.ttf".
    /// </summary>
    [TestClass]
    public class OtsSubsetMatrixTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        // Each line targets a specific part of the subsetter
        private const string ChallengingText =
            "office affluent final ffi ffl fj " +                          // GSUB ligatures
            "AVATAR Wave To Yo LT P. F. W. " +                            // GPOS kerning (PairPos format 1 and 2)
            "Åsa Öberg äter smörgåsar på Café Résumé, naïve façade ñ " +  // precomposed diacritics, composite glyphs
            "e\u0301 a\u0308 o\u0302\u0301 n\u0303 " +                    // combining marks (MarkToBase, MarkToMark)
            "Tiếng Việt: Phở, đường, ổn định " +                          // stacked diacritics
            "½ ¼ € 1 234,56 −12 % ‰ № “citat” ‘enkel’ – — … • " +        // figures and punctuation
            "Ελληνικά Кириллица " +                                       // Greek and Cyrillic
            "\u2764\uFE0F \U0001F600 \U0001D400";                        // variation selector, emoji, math bold A (cmap 12/14)

        [TestMethod]
        [DataRow("Mulish-Regular.ttf")]
        [DataRow("Mulish-Regular-GsubExtension.ttf")]
        [DataRow("Roboto-Regular.ttf")]
        [DataRow("OpenSans-Regular.ttf")]
        [DataRow("EBGaramond-Regular.ttf")]
        [DataRow("CrimsonText-Regular.ttf")]
        public void OtsMatrix_TestFolderFonts(string fileName)
        {
            CreateAndSaveSubset(Path.Combine(FontFolder, fileName));
        }

        [TestMethod]
        [DataRow("calibri.ttf")]
        [DataRow("arial.ttf")]
        [DataRow("times.ttf")]
        [DataRow("segoeui.ttf")]
        [DataRow("georgia.ttf")]
        [DataRow("seguisym.ttf")]
        public void OtsMatrix_SystemFonts(string fileName)
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), fileName);
            if (!File.Exists(path))
            {
                Assert.Inconclusive(fileName + " is not installed on this machine");
            }
            CreateAndSaveSubset(path);
        }

        private void CreateAndSaveSubset(string fontPath)
        {
            var font = new OpenTypeFont(File.ReadAllBytes(fontPath));
            var subsetFont = font.CreateSubset(ChallengingText);

            var bytes = new OpenTypeFontSerializer(subsetFont).Serialize();
            var parsedFont = new OpenTypeFont(bytes);

            // Save before asserting, so the file is available for OTS even if an assert fails
            SaveFontForCurrentTest(parsedFont, Path.GetFileNameWithoutExtension(fontPath));

            Console.WriteLine("Glyphs: " + parsedFont.MaxpTable.numGlyphs);
            Console.WriteLine("Tables: " + string.Join(", ", parsedFont.TableRecords.Keys.ToArray()));
            Console.WriteLine("GSUB lookup types: " + string.Join(", ",
                FontTestHelper.GetGsubLookupTypes(parsedFont).Select(t => t.ToString()).ToArray()));

            FontTestHelper.AssertGsubLookupIndicesValid(parsedFont);
        }
    }
}