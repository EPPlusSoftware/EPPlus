using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class GlyphTableValidationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void LocaTableValidation_Test()
        {
            var font = OpenTypeFonts.LoadFont("Roboto");
            var validator = new LocaTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.LocaTable, context);
            Assert.IsTrue(result.IsValid, "Loca validation failed for a known good font.");
        }

        [TestMethod]
        public void HmtxTableValidation_Test()
        {
            var font = OpenTypeFonts.LoadFont("Roboto");
            var validator = new HmtxTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HmtxTable, context);
            Assert.IsTrue(result.IsValid, "Hmtx validation failed for a known good font.");
        }

        [TestMethod]
        [DataRow("Roboto", FontSubFamily.Regular)]
        [DataRow("Roboto", FontSubFamily.Italic)]
        [DataRow("EB Garamond", FontSubFamily.Regular)]
        [DataRow("Mulish", FontSubFamily.Regular)]
        public void GlyfTableValidation_Test(string fontName, FontSubFamily subFamily)
        {
            var font = OpenTypeFonts.LoadFont(fontName, subFamily);
            var validator = new GlyfTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.GlyfTable, context);
            Assert.IsTrue(result.IsValid, $"Glyf validation failed for a known good font: {fontName} {subFamily}");
        }
    }
}
