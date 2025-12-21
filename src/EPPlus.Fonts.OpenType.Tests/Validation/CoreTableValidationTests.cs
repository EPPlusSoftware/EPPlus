using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    [TestClass]
    public class CoreTableValidationTests : ValidationTestBase
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            ValidationTestHelper.ClassInitialize(testContext);
        }

        [TestMethod]
        public void HeadTable_Validation_ShouldPass()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new HeadTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HeadTable, context);
            Assert.IsTrue(result.IsValid);
        }


        [TestMethod]
        public void MaxpTable_Validation_ShouldPass()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new MaxpTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.MaxpTable, context);
            Assert.IsTrue(result.IsValid, "MaxpTable validation failed for a known good font.");
        }

        [TestMethod]
        public void HheaTable_Validation_ShouldPass()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new HheaTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HheaTable, context);
            Assert.IsTrue(result.IsValid, "Hhea validation failed for a known good font.");
        }

        [TestMethod]
        public void NameTable_Validation_ShouldPass()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new NameTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.NameTable, context);
            Assert.IsTrue(result.IsValid, "Name validation failed for a known good font.");
        }

        [TestMethod]
        public void Os2Table_Validation_ShouldPass()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new Os2TableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.Os2Table, context);
            Assert.IsTrue(result.IsValid, "Os/2 validation failed for a known good font.");
        }

        [TestMethod]
        public void PostTable_Validation_ShouldPass()
        {
            var font = OpenTypeFonts.GetFontData(FontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new PostTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.PostTable, context);
            Assert.IsTrue(result.IsValid, "Post validation failed for a known good font.");
        }
    }
}
