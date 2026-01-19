using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class EntireFontValidationTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
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
