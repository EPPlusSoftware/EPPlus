using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tests.Helpers;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class GsubTableValidationTests : FontTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            FontDirectoriesTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void GsubTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new GsubTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.GsubTable, context);
            Assert.IsTrue(result.IsValid, "Gsub validation failed for a known good font.");
        }
    }
}
