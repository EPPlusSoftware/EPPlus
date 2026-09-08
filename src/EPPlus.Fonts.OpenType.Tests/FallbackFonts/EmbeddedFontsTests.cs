using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.FallbackFonts
{
    [TestClass]
    public class EmbeddedFontsTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void BundledFamilies_MatchesEmbeddedResources()
        {
            Assert.IsTrue(EmbeddedFonts.IsBundledFamily(
                EmbeddedFonts.LoadNotoEmoji().GetEnglishFontFamilyName()));
            Assert.IsTrue(EmbeddedFonts.IsBundledFamily(
                EmbeddedFonts.LoadNotoMath().GetEnglishFontFamilyName()));
            foreach (FontSubFamily sf in Enum.GetValues(typeof(FontSubFamily)))
            {
                Assert.IsTrue(EmbeddedFonts.IsBundledFamily(
                    EmbeddedFonts.LoadArchivoNarrow(sf).GetEnglishFontFamilyName()), sf.ToString());
            }
        }
    }
}
