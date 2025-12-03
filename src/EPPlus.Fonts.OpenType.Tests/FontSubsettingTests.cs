using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Head;
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
        public void TestHeadTable()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var subsetFont = font.CreateSubset(new char[] { 'a', 'b', 'c' });
            var validator = new HeadTableValidator();
            var isValid = validator.Validate(subsetFont.HeadTable, new FontValidationContext(subsetFont));
            var bytes = subsetFont.Serialize();
            // read the font back
            var reader = new FontsBinaryReader(new MemoryStream(bytes));
            var parsedFont = new OpenTypeFont(reader, font.Format);
            var isParsedFontValid = parsedFont.ValidateFont(FontValidationSeverity.Warning);
            Assert.AreEqual(1, parsedFont.TableRecords.Count);
        }


        [TestMethod]
        public void TestHeadAndName_SubsetSerializationRoundtrip()
        {
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Add name table in builder (via Clone)
            Assert.IsNotNull(subsetFont.NameTable);

            // Preprocess must have been called by your CreateSubset flow already
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();

            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);
            Assert.AreEqual(2, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));

            // Validate both tables
            var headValid = new HeadTableValidator().Validate(parsedFont.HeadTable, new FontValidationContext(parsedFont));
            var nameValid = new NameTableValidator().Validate(parsedFont.NameTable, new FontValidationContext(parsedFont));

            Assert.IsTrue(headValid.IsValid);
            Assert.IsTrue(nameValid.IsValid);
        }


        [TestMethod]
        public void TestHeadNameAndMaxp_SubsetSerializationRoundtrip()
        {
            // Arrange
            var font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            var subsetFont = font.CreateSubset(new[] { 'a', 'b', 'c' });

            // Act
            var serializer = new OpenTypeFontSerializer(subsetFont);
            var bytes = serializer.Serialize();

            var parsedFont = new OpenTypeFont(new FontsBinaryReader(new MemoryStream(bytes)), font.Format);

            // Assert: Check table count and presence
            Assert.AreEqual(3, parsedFont.TableRecords.Count);
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("head"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("name"));
            Assert.IsTrue(parsedFont.TableRecords.ContainsKey("maxp"));

            // Validate head
            var headValid = new HeadTableValidator().Validate(parsedFont.HeadTable, new FontValidationContext(parsedFont));
            Assert.IsTrue(headValid.IsValid);

            // Validate name
            var nameValid = new NameTableValidator().Validate(parsedFont.NameTable, new FontValidationContext(parsedFont));
            Assert.IsTrue(nameValid.IsValid);

            // Validate maxp
            var maxpValid = new MaxpTableValidator().Validate(parsedFont.MaxpTable, new FontValidationContext(parsedFont));
            Assert.IsTrue(maxpValid.IsValid);

            // Extra check: numGlyphs should match glyphSet.Count
            var glyphSet = new HashSet<ushort>();
            foreach (var ch in new[] { 'a', 'b', 'c' })
            {
                if (font.CmapTable.TryGetGlyphId(ch, out ushort glyphId))
                    glyphSet.Add(glyphId);
            }
            glyphSet.Add(0); // .notdef

            Assert.AreEqual((ushort)glyphSet.Count, parsedFont.MaxpTable.numGlyphs);
        }

    }
}
