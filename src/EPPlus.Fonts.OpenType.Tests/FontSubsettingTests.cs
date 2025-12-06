using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
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
