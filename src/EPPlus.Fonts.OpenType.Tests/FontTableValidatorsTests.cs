using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Gsub;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using Microsoft.Testing.Platform.Logging;
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
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var report = font.ValidateFont(FontValidationSeverity.Error);
            Assert.IsTrue(report.IsValid);
        }

        [TestMethod]
        public void HeadTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new HeadTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HeadTable, context);
            Assert.IsTrue(result.IsValid);
        }


        [TestMethod]
        public void MaxpTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new MaxpTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.MaxpTable, context);
            Assert.IsTrue(result.IsValid, "MaxpTable validation failed for a known good font.");
        }

        [TestMethod]
        public void HheaTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new HheaTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HheaTable, context);
            Assert.IsTrue(result.IsValid, "Hhea validation failed for a known good font.");
        }

        [TestMethod]
        public void NameTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new NameTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.NameTable, context);
            Assert.IsTrue(result.IsValid, "Name validation failed for a known good font.");
        }

        [TestMethod]
        public void Os2TableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new Os2TableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.Os2Table, context);
            Assert.IsTrue(result.IsValid, "Os/2 validation failed for a known good font.");
        }

        [TestMethod]
        public void PostTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new PostTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.PostTable, context);
            Assert.IsTrue(result.IsValid, "Post validation failed for a known good font.");
        }

        [TestMethod]
        public void CmapTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new CmapTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.CmapTable, context);
            Assert.IsTrue(result.IsValid, "Cmap validation failed for a known good font.");
        }

        [TestMethod]
        public void LocaTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new LocaTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.LocaTable, context);
            Assert.IsTrue(result.IsValid, "Loca validation failed for a known good font.");
        }

        [TestMethod]
        public void HmtxTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new HmtxTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.HmtxTable, context);
            Assert.IsTrue(result.IsValid, "Hmtx validation failed for a known good font.");
        }

        [TestMethod]
        public void GsubTableValidation_Test()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false, true);
            var validator = new GsubTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.GsubTable, context);
            Assert.IsTrue(result.IsValid, "Gsub validation failed for a known good font.");
        }

        [TestMethod]
        [DataRow("Roboto", FontSubFamily.Regular)]
        [DataRow("Roboto", FontSubFamily.Italic)]
        [DataRow("EB Garamond", FontSubFamily.Regular)]
        [DataRow("Mulish", FontSubFamily.Regular)]
        public void GlyfTableValidation_Test(string fontName, FontSubFamily subFamily)
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, fontName, subFamily, false, true);
            var validator = new GlyfTableValidator();
            var context = new FontValidationContext(font);
            var result = validator.Validate(font.GlyfTable, context);
            Assert.IsTrue(result.IsValid, $"Glyf validation failed for a known good font: {fontName} {subFamily}");
        }
    }
}
