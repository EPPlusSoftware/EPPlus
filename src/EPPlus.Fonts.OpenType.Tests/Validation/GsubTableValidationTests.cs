using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class GsubTableValidationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        [DataRow("Roboto")]
        [DataRow("OpenSans")]
        [DataRow("SourceSans3")]
        [DataRow("NotoEmoji")]
        public void GsubTableValidation_Test(string fontName)
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, fontName, FontSubFamily.Regular, false, true);
            var validator = new GsubTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.GsubTable, context);
            Assert.IsTrue(result.IsValid, $"Gsub validation failed for a known good font. {fontName}");
        }
    }
}
