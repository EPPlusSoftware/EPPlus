using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class CmapTableValidationTests : ValidationTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            ValidationTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void CmapTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new CmapTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.CmapTable, context);
            Assert.IsTrue(result.IsValid, "Cmap validation failed for a known good font.");
        }
    }
}
