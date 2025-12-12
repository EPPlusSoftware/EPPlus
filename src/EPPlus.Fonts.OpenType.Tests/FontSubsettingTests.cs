using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class FontSubsettingTests
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
        public void TestHeadNameMaxpHheaLocaGlyfAndCmap_SubsetSerializationRoundtrip()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Assert: Check table count and presence
            Assert.AreEqual(10, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hhea"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("loca"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("glyf"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("cmap"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("post"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("OS/2"));

            // Validate tables
            Assert.IsTrue(new HeadTableValidator().Validate(parsedFont.HeadTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new NameTableValidator().Validate(parsedFont.NameTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new MaxpTableValidator().Validate(parsedFont.MaxpTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new HheaTableValidator().Validate(parsedFont.HheaTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new GlyfTableValidator().Validate(parsedFont.GlyfTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new LocaTableValidator().Validate(parsedFont.LocaTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new CmapTableValidator().Validate(parsedFont.CmapTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new PostTableValidator().Validate(parsedFont.PostTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new Os2TableValidator().Validate(parsedFont.Os2Table, new FontValidationContext(parsedFont)).IsValid);

            // Extra check: glyphSet.Count should match numGlyphs and numberOfHMetrics
            var glyphSet = new HashSet<ushort>();
            foreach (var ch in new[] { 'a', 'b', 'c' })
            {
                if (font.CmapTable.TryGetGlyphId(ch, out ushort glyphId))
                    glyphSet.Add(glyphId);
            }
            glyphSet.Add(0); // Always include .notdef

            // Extra check: glyph count should be requested glyphs + space + .notdef
            var requestedChars = new[] { 'a', 'b', 'c' };
            int expectedGlyphs = requestedChars.Length;     // 3
            expectedGlyphs += 1;                            // + space (U+0020) – always included
            expectedGlyphs += 1;                            // + .notdef (GID 0)

            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.MaxpTable.numGlyphs);     // 5
            Assert.AreEqual((ushort)expectedGlyphs, parsedFont.HheaTable.numberOfHMetrics); // 5

            // Bonus: verify that space actually exists in cmap
            Assert.IsTrue(parsedFont.CmapTable.ContainsChar(32));
        }

        [TestMethod]
        public void TestSubsetWithFullValidation()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            //var path = @"c:\Temp\subset_font.ttf";
            //if(File.Exists(path)) File.Delete(path);
            //File.WriteAllBytes(@"c:\Temp\subset_font.ttf", bytes);

            var report = new FontValidator().Validate(parsedFont, FontValidationSeverity.Warning);
            Assert.IsTrue(report.IsValid);
        }

        [TestMethod]
        public void TestSubsetWithFullValidation1()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Mulish", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a',});

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            //var path = @"c:\Temp\subset_font.ttf";
            //if (File.Exists(path)) File.Delete(path);
            //File.WriteAllBytes(@"c:\Temp\subset_font.ttf", bytes);

            var report = new FontValidator().Validate(parsedFont, FontValidationSeverity.Warning);
            Assert.IsTrue(report.IsValid);
        }

        [TestMethod]
        public void TestSubsetWithFullValidation2()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'F', 'l', 'y', 'g', 'a', 'n', 'd', 'e', 'b', 'ä', 'c', 'k', 's', 'i', 'r', 'ö', 'h', 'w', 'p', 'å', 'm', 'j', 'u', 't', 'v', 'o', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            //var path = @"c:\Temp\subset_font2.ttf";
            //if (File.Exists(path)) File.Delete(path);
            //File.WriteAllBytes(@"c:\Temp\subset_font2.ttf", bytes);

            var report = new FontValidator().Validate(parsedFont, FontValidationSeverity.Warning);
            Assert.IsTrue(report.IsValid);
        }

        [TestMethod]
        public void Subset_Roboto_With_ÅÄÖ_Should_Work()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular);

            // get the original å
            var ågId = font.CmapTable.MapCharToGlyph('å');
            var åglyph = font.GlyfTable.GetGlyph((ushort)ågId);

            var subset = font.CreateSubset("Testar åäö ÅÄÖ och även é û č ć đ ł".Distinct());

            // Save for inspection (optional)
            //File.WriteAllBytes(@"C:\temp\Roboto-subset-aao.ttf", subset.Serialize());

            // Verify that 'å' actually has a composite glyph
            var åGlyphId = subset.CmapTable.MapCharToGlyph('å');
            var glyph = subset.GlyfTable.GetGlyph((ushort)åGlyphId);

            Assert.IsTrue(glyph.Header.numberOfContours < 0); // måste vara composite
            Assert.IsTrue(glyph.CompositeData.Components.Count > 0);
        }

        [TestMethod]
        public void Subset_Mulish_With_ÅÄÖ_Should_Work()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Mulish", FontSubFamily.Regular);

            var subset = font.CreateSubset("Testar åäö ÅÄÖ och även é û č ć đ ł".Distinct());

            // Save for inspection (optional)
            //File.WriteAllBytes(@"C:\temp\Mulish-subset-aao.ttf", subset.Serialize());

            // Verify that 'å' actually has a composite glyph
            var åGlyphId = subset.CmapTable.MapCharToGlyph('å');
            var glyph = subset.GlyfTable.GetGlyph((ushort)åGlyphId);

            Assert.IsTrue(glyph.Header.numberOfContours < 0); // måste vara composite
            Assert.IsTrue(glyph.CompositeData.Components.Count > 0);
        }

        [TestMethod]
        public void Subset_BIZUDGothic_With_ÅÄÖ_Should_Work()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "BIZUDGothic", FontSubFamily.Regular);
            var cmt = font.CmapTable;
            var subset = font.CreateSubset("Testar åäö ÅÄÖ och även é û č ć đ ł".Distinct());

            // Save for inspection (optional)
            //File.WriteAllBytes(@"C:\temp\BIZUDGothic-subset-aao.ttf", subset.Serialize());

            // Verify that 'å' actually has a composite glyph
            var åGlyphId = subset.CmapTable.MapCharToGlyph('å');
            var glyph = subset.GlyfTable.GetGlyph((ushort)åGlyphId);

            Assert.IsTrue(glyph.Header.numberOfContours == 4); // måste vara composite
            Assert.IsNotNull(glyph.SimpleData);
            Assert.IsNull(glyph.CompositeData);
        }

    }
}
