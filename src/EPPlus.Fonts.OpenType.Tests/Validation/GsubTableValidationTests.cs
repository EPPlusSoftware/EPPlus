using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class GsubTableValidationTests : ValidationTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            ValidationTestHelper.ClassInitialize(testContext);
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
