using EPPlus.Fonts.OpenType.FontValidation;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class EntireFontValidationTests : ValidationTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            ValidationTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void ValidateEntireFont()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var report = font.ValidateFont(FontValidationSeverity.Error);
            Assert.IsTrue(report.IsValid);
        }
    }
}
