/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/27/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.FallbackFonts
{
    [TestClass]
    public class DefaultPrimaryFontTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private static readonly string UnknownFont = "ThisFontDoesNotExist_XYZ";

        private static DefaultFontResolver CreateIsolatedResolver()
        {
            // No directories, no system fonts — guarantees Archivo Narrow fallback
            return new DefaultFontResolver(fontDirectories: null, searchSystemDirectories: false);
        }

        [TestMethod]
        public void ResolveFont_UnknownFont_Regular_ShouldFallbackToArchivoNarrow()
        {
            var resolver = CreateIsolatedResolver();
            var bytes = resolver.ResolveFont(UnknownFont, FontSubFamily.Regular);

            Assert.IsNotNull(bytes, "Should fall back to Archivo Narrow, not return null");
            var font = OpenTypeFonts.GetFromBytes(bytes);
            Assert.AreEqual("Archivo Narrow", font.NameTable.GetFamilyName());
            Assert.AreEqual(FontSubFamily.Regular, font.NameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void ResolveFont_UnknownFont_Bold_ShouldFallbackToArchivoNarrowBold()
        {
            var resolver = CreateIsolatedResolver();
            var bytes = resolver.ResolveFont(UnknownFont, FontSubFamily.Bold);

            Assert.IsNotNull(bytes, "Should fall back to Archivo Narrow Bold, not return null");
            var font = OpenTypeFonts.GetFromBytes(bytes);
            Assert.AreEqual("Archivo Narrow", font.NameTable.GetFamilyName());
            Assert.AreEqual(FontSubFamily.Bold, font.NameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void ResolveFont_UnknownFont_Italic_ShouldFallbackToArchivoNarrowItalic()
        {
            var resolver = CreateIsolatedResolver();
            var bytes = resolver.ResolveFont(UnknownFont, FontSubFamily.Italic);

            Assert.IsNotNull(bytes, "Should fall back to Archivo Narrow Italic, not return null");
            var font = OpenTypeFonts.GetFromBytes(bytes);
            Assert.AreEqual("Archivo Narrow", font.NameTable.GetFamilyName());
            Assert.AreEqual(FontSubFamily.Italic, font.NameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void ResolveFont_UnknownFont_BoldItalic_ShouldFallbackToArchivoNarrowBoldItalic()
        {
            var resolver = CreateIsolatedResolver();
            var bytes = resolver.ResolveFont(UnknownFont, FontSubFamily.BoldItalic);

            Assert.IsNotNull(bytes, "Should fall back to Archivo Narrow Bold Italic, not return null");
            var font = OpenTypeFonts.GetFromBytes(bytes);
            Assert.AreEqual("Archivo Narrow", font.NameTable.GetFamilyName());
            Assert.AreEqual(FontSubFamily.BoldItalic, font.NameTable.GetSubfamilyEnum());
        }
    }
}