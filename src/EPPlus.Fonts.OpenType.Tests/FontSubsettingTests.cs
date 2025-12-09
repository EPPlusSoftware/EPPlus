using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
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
        public void TestHeadNameMaxpAndHhea_SubsetSerializationRoundtrip()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Assert: Check table count and presence
            Assert.AreEqual(5, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hhea"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));

            // Validate tables
            Assert.IsTrue(new HeadTableValidator().Validate(parsedFont.HeadTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new NameTableValidator().Validate(parsedFont.NameTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new MaxpTableValidator().Validate(parsedFont.MaxpTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new HheaTableValidator().Validate(parsedFont.HheaTable, new FontValidationContext(parsedFont)).IsValid);

            // Extra check: glyphSet.Count should match numGlyphs and numberOfHMetrics
            var glyphSet = new HashSet<ushort>();
            foreach (var ch in new[] { 'a', 'b', 'c' })
            {
                if (font.CmapTable.TryGetGlyphId(ch, out ushort glyphId))
                    glyphSet.Add(glyphId);
            }
            glyphSet.Add(0); // Always include .notdef

            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.MaxpTable.numGlyphs);
            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.HheaTable.numberOfHMetrics);
        }

        [TestMethod]
        public void TestHeadNameMaxpHheaLocaAndGlyf_SubsetSerializationRoundtrip()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Assert: Check table count and presence
            Assert.AreEqual(7, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hhea"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("loca"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("glyf"));

            // Validate tables
            Assert.IsTrue(new HeadTableValidator().Validate(parsedFont.HeadTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new NameTableValidator().Validate(parsedFont.NameTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new MaxpTableValidator().Validate(parsedFont.MaxpTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new HheaTableValidator().Validate(parsedFont.HheaTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new GlyfTableValidator().Validate(parsedFont.GlyfTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new LocaTableValidator().Validate(parsedFont.LocaTable, new FontValidationContext(parsedFont)).IsValid);

            // Extra check: glyphSet.Count should match numGlyphs and numberOfHMetrics
            var glyphSet = new HashSet<ushort>();
            foreach (var ch in new[] { 'a', 'b', 'c' })
            {
                if (font.CmapTable.TryGetGlyphId(ch, out ushort glyphId))
                    glyphSet.Add(glyphId);
            }
            glyphSet.Add(0); // Always include .notdef

            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.MaxpTable.numGlyphs);
            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.HheaTable.numberOfHMetrics);
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
            Assert.AreEqual(8, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hhea"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("loca"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("glyf"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("cmap"));

            // Validate tables
            Assert.IsTrue(new HeadTableValidator().Validate(parsedFont.HeadTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new NameTableValidator().Validate(parsedFont.NameTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new MaxpTableValidator().Validate(parsedFont.MaxpTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new HheaTableValidator().Validate(parsedFont.HheaTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new GlyfTableValidator().Validate(parsedFont.GlyfTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new LocaTableValidator().Validate(parsedFont.LocaTable, new FontValidationContext(parsedFont)).IsValid);
            Assert.IsTrue(new CmapTableValidator().Validate(parsedFont.CmapTable, new FontValidationContext(parsedFont)).IsValid);

            // Extra check: glyphSet.Count should match numGlyphs and numberOfHMetrics
            var glyphSet = new HashSet<ushort>();
            foreach (var ch in new[] { 'a', 'b', 'c' })
            {
                if (font.CmapTable.TryGetGlyphId(ch, out ushort glyphId))
                    glyphSet.Add(glyphId);
            }
            glyphSet.Add(0); // Always include .notdef

            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.MaxpTable.numGlyphs);
            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.HheaTable.numberOfHMetrics);
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

            var path = @"c:\Temp\subset_font.ttf";
            if(File.Exists(path)) File.Delete(path);
            File.WriteAllBytes(@"c:\Temp\subset_font.ttf", bytes);

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

            var path = @"c:\Temp\subset_font.ttf";
            if (File.Exists(path)) File.Delete(path);
            File.WriteAllBytes(@"c:\Temp\subset_font.ttf", bytes);

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

            var path = @"c:\Temp\subset_font2.ttf";
            if (File.Exists(path)) File.Delete(path);
            File.WriteAllBytes(@"c:\Temp\subset_font2.ttf", bytes);

            var report = new FontValidator().Validate(parsedFont, FontValidationSeverity.Warning);
            Assert.IsTrue(report.IsValid);
        }


        [TestMethod]
        public void TestHeadNameMaxpHheaAndHmtx_SubsetSerializationRoundtrip()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", FontSubFamily.Regular, false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();
            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            Assert.AreEqual(5, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("hmtx"));

            var hmtxValid = new HmtxTableValidator().Validate(parsedFont.HmtxTable, new FontValidationContext(parsedFont));
            Assert.IsTrue(hmtxValid.IsValid);

            Assert.AreEqual(parsedFont.HheaTable.numberOfHMetrics, parsedFont.HmtxTable.hMetrics.Count);
        }

    }
}
