using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class FontTableValidatorsTests
    {
        private static string _fontFolder = string.Empty;
        private static List<string> _fontFolders = new List<string>();

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
            _fontFolders.Clear();
            _fontFolders.Add(_fontFolder);
            OpenTypeFonts.ClearFontCache();
        }

        [TestMethod]
        public void ValidateEntireFont()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var report = font.ValidateFont();
            Assert.IsTrue(report.IsValid);
        }

        [TestMethod]
        public void HeadTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var validator = new HeadTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HeadTable, context);
            Assert.IsTrue(result.IsValid);
        }


        [TestMethod]
        public void MaxpTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var validator = new MaxpTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.MaxpTable, context);
            Assert.IsTrue(result.IsValid, "MaxpTable validation failed for a known good font.");
        }

        [TestMethod]
        public void HheaTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var validator = new HheaTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HheaTable, context);
            Assert.IsTrue(result.IsValid, "Hhea validation failed for a known good font.");
        }

        [TestMethod]
        public void NameTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var validator = new NameTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.NameTable, context);
            Assert.IsTrue(result.IsValid, "Name validation failed for a known good font.");
        }

        [TestMethod]
        public void Os2TableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var validator = new Os2TableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.Os2Table, context);
            Assert.IsTrue(result.IsValid, "Os/2 validation failed for a known good font.");
        }
    }
}
