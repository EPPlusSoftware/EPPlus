using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class EntireFontValidationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void ValidateEntireFont()
        {
            var font = OpenTypeFonts.LoadFont("Roboto");
            var report = font.ValidateFont(FontValidationSeverity.Error);
            Assert.IsTrue(report.IsValid);
        }
    }
}
