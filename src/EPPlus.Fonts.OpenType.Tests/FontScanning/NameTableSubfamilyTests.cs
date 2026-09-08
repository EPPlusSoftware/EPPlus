/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/26/2026         EPPlus Software AB           Initial tests for NameTable family/subfamily naming
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontLocalization;
using EPPlus.Fonts.OpenType.Tables.Name;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.FontScanning
{
    /// <summary>
    /// Unit tests for how NameTable resolves family and subfamily names — bare table instances,
    /// no font file loading, so they stay fast and deterministic regardless of installed fonts.
    ///
    /// THE PAIR PRINCIPLE (what most of these tests defend)
    /// -----------------------------------------------------
    /// OpenType fonts can carry two parallel, complete naming systems:
    ///
    ///   legacy / RIBBI    nameID 1  (family) + nameID 2  (subfamily)
    ///   typographic       nameID 16 (family) + nameID 17 (subfamily)
    ///
    /// Both describe the same file correctly, but they are PAIRS and must never be mixed.
    /// Arial Black (ariblk.ttf) reads:
    ///
    ///   nameID 1  = "Arial Black"     nameID 2  = "Regular"
    ///   nameID 16 = "Arial"           nameID 17 = "Black"
    ///
    /// Taking the family from one system and the subfamily from the other yields the pair
    /// "Arial Black" + "Black", which exists in neither system. That was the original bug:
    /// a request for "Arial Black" + Regular matched the family but not the style, so
    /// IsExactMatch was false and DefaultFontResolver fell through to the built-in fallback
    /// chain ("Arial Black" -> "Liberation Sans" -> "Arial") and returned plain Arial, even
    /// though Arial Black was installed.
    ///
    /// EPPlus uses the legacy/RIBBI system, because FontSubFamily's four values
    /// (Regular/Bold/Italic/BoldItalic) ARE the RIBBI model, and nameID 2 is guaranteed by
    /// the spec to be one of those four. It is also the view Windows, GDI and Excel present.
    /// </summary>
    [TestClass]
    public class NameTableSubfamilyTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region The pair principle — real ariblk.ttf field layout

        /// <summary>
        /// The exact four-field layout found in C:\Windows\Fonts\ariblk.ttf. This is the
        /// regression test for the original bug and the single most important test here.
        /// </summary>
        [TestMethod]
        public void ArialBlackLayout_ResolvesToLegacyPair_NotAMixOfBothSystems()
        {
            var nameTable = CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.FontFamilyName, "Arial Black"),
                MakeEnglishRecord(NameRecordTypes.FontSubfamilyName, "Regular"),
                MakeEnglishRecord(NameRecordTypes.TypographicFamilyName, "Arial"),
                MakeEnglishRecord(NameRecordTypes.TypographicSubfamilyName, "Black"));

            var family = nameTable.GetFamilyName();
            var subfamily = nameTable.GetSubfamilyName();

            // Both halves must come from the SAME system. Asserting them together (rather than
            // in two separate tests) is deliberate: the failure mode is a mismatched pair, and
            // either value on its own looks perfectly reasonable.
            Assert.AreEqual("Arial Black", family,
                "Family must come from nameID 1 (legacy), not nameID 16 ('Arial').");
            Assert.AreEqual("Regular", subfamily,
                "Subfamily must come from nameID 2 (legacy), not nameID 17 ('Black'). " +
                "Reading nameID 17 here produces the impossible pair 'Arial Black' + 'Black'.");
            Assert.AreEqual(FontSubFamily.Regular, nameTable.GetSubfamilyEnum());
        }

        /// <summary>
        /// Guard against someone "fixing" GetSubfamilyName to prefer the newer nameID 17.
        /// Deliberately makes the two systems disagree so preferring 17 is unmistakable.
        /// </summary>
        [TestMethod]
        public void GetSubfamilyName_DoesNotPreferTypographicSubfamily17()
        {
            var nameTable = CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.FontSubfamilyName, "Regular"),
                MakeEnglishRecord(NameRecordTypes.TypographicSubfamilyName, "Black"));

            Assert.AreEqual("Regular", nameTable.GetSubfamilyName(),
                "nameID 17 must not win over nameID 2. nameID 17 belongs to the typographic " +
                "system (paired with nameID 16) and carries weights outside the RIBBI model.");
        }

        /// <summary>
        /// Mirror of the above for the family side — guards the ID1-over-ID16 priority that
        /// GetFamilyName's (previously contradictory) doc comment used to describe backwards.
        /// </summary>
        [TestMethod]
        public void GetFamilyName_DoesNotPreferTypographicFamily16()
        {
            var nameTable = CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.FontFamilyName, "Arial Black"),
                MakeEnglishRecord(NameRecordTypes.TypographicFamilyName, "Arial"));

            Assert.AreEqual("Arial Black", nameTable.GetFamilyName(),
                "nameID 16 must not win over nameID 1, or 'Arial Black' collapses into the " +
                "'Arial' family and can no longer be resolved as a distinct font.");
        }

        #endregion

        #region Typographic system as last resort — only when the legacy field is absent

        [TestMethod]
        public void GetSubfamilyName_NoNameId2_FallsBackToTypographicSubfamily17()
        {
            // A font that omits nameID 2 entirely. Then nameID 17 is all we have, and using
            // it is correct — the pair principle only forbids mixing when BOTH are present.
            var nameTable = CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.TypographicSubfamilyName, "Bold"));

            Assert.AreEqual("Bold", nameTable.GetSubfamilyName());
            Assert.AreEqual(FontSubFamily.Bold, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetFamilyName_NoNameId1_FallsBackToTypographicFamily16()
        {
            var nameTable = CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.TypographicFamilyName, "Arial"));

            Assert.AreEqual("Arial", nameTable.GetFamilyName());
        }

        #endregion

        #region English must win over localized records

        /// <summary>
        /// ariblk.ttf carries 75 name records, including a dozen localized nameID 2 values
        /// ("Normal", "obycejne", "Standard", "Kanonika", "Obychnyy", "Arrunta", ...).
        /// Picking whichever comes first in file order happens to work for ariblk.ttf, but
        /// that is luck, not a guarantee — file order is entirely up to the font vendor.
        /// </summary>
        [TestMethod]
        public void GetSubfamilyName_LocalizedRecordFirst_StillPrefersEnglish()
        {
            var nameTable = CreateNameTable(
                MakeLocalizedRecord(NameRecordTypes.FontSubfamilyName, "Fet"),      // sv-SE
                MakeEnglishRecord(NameRecordTypes.FontSubfamilyName, "Bold"));

            Assert.AreEqual("Bold", nameTable.GetSubfamilyName(),
                "A localized nameID 2 appearing earlier in the table must not beat the " +
                "English one, or the subfamily string becomes unparseable by the enum mapping.");
            Assert.AreEqual(FontSubFamily.Bold, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetFamilyName_LocalizedRecordFirst_StillPrefersEnglish()
        {
            var nameTable = CreateNameTable(
                MakeLocalizedRecord(NameRecordTypes.FontFamilyName, "Arial Svart"),
                MakeEnglishRecord(NameRecordTypes.FontFamilyName, "Arial Black"));

            Assert.AreEqual("Arial Black", nameTable.GetFamilyName());
        }

        [TestMethod]
        public void GetSubfamilyName_OnlyLocalizedAvailable_UsesItRatherThanNothing()
        {
            // No English record at all — better to return the localized string than to fall
            // through to the typographic system or the "Regular" default.
            var nameTable = CreateNameTable(
                MakeLocalizedRecord(NameRecordTypes.FontSubfamilyName, "Normal"));

            Assert.AreEqual("Normal", nameTable.GetSubfamilyName());
            Assert.AreEqual(FontSubFamily.Regular, nameTable.GetSubfamilyEnum(),
                "'Normal' is one of the recognized Regular spellings.");
        }

        #endregion

        #region GetSubfamilyEnum — RIBBI mapping (regression guards)

        [TestMethod]
        public void GetSubfamilyEnum_Regular_ReturnsRegular()
        {
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Regular"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_Bold_ReturnsBold()
        {
            Assert.AreEqual(FontSubFamily.Bold, EnumFor("Bold"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_Italic_ReturnsItalic()
        {
            Assert.AreEqual(FontSubFamily.Italic, EnumFor("Italic"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_BoldItalic_ReturnsBoldItalic()
        {
            Assert.AreEqual(FontSubFamily.BoldItalic, EnumFor("Bold Italic"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_Oblique_ReturnsItalic()
        {
            Assert.AreEqual(FontSubFamily.Italic, EnumFor("Oblique"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_AlternateRegularSpellings_ReturnRegular()
        {
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Normal"));
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Roman"));
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Book"));
        }

        #endregion

        #region GetSubfamilyEnum — weight names beyond Bold (defense in depth)

        // With the nameID 2 priority fixed, a well-formed font never reaches these branches:
        // nameID 2 is always a RIBBI name. They still matter for fonts that omit nameID 2 and
        // fall back to nameID 17, which is where weights like "Black" or "Light" show up.

        [TestMethod]
        public void GetSubfamilyEnum_WeightNamesBeyondBold_ReturnRegularNotBold()
        {
            // These are separate typographic weights, already distinguished by the family
            // name (e.g. "Arial Black"). Within the 4-value enum their base instance is
            // Regular. Mapping them to Bold would disqualify an exact Regular match, and
            // would let a Bold request be satisfied by a far heavier face than intended.
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Black"), "Black");
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Heavy"), "Heavy");
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Demi"), "Demi");
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Light"), "Light");
            Assert.AreEqual(FontSubFamily.Regular, EnumFor("Medium"), "Medium");
        }

        [TestMethod]
        public void GetSubfamilyEnum_BlackItalic_ReturnsItalic()
        {
            // Bodoni MT Black and Segoe UI Black both ship a "Black Italic" face. The weight
            // is dropped (no enum value for it) but the italic axis is real and must survive.
            Assert.AreEqual(FontSubFamily.Italic, EnumFor("Black Italic"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_SemiBold_ReturnsBold()
        {
            // "Semibold" legitimately contains the substring "bold", unlike Black/Heavy/Demi,
            // and Bold is the closest of the four values. This behaviour is intentional.
            Assert.AreEqual(FontSubFamily.Bold, EnumFor("SemiBold"));
        }

        [TestMethod]
        public void GetSubfamilyEnum_WeightNameBeyondBold_DoesNotConsultFsSelection()
        {
            // Regression guard for a subtle trap: if the weight-name branch falls through to
            // the OS/2 fsSelection fallback instead of returning Regular explicitly, the bug
            // reappears through a different path. Vendors commonly set fsSelection's BOLD bit
            // on Black/Heavy faces as a legacy hint for apps that can't read the name table.
            const ushort fsSelectionBold = 0x0020;
            var nameTable = CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.TypographicSubfamilyName, "Black"));
            nameTable.Os2FsSelection = fsSelectionBold;

            Assert.AreEqual(FontSubFamily.Regular, nameTable.GetSubfamilyEnum(),
                "The name table gave a usable answer, so fsSelection must not be consulted.");
        }

        #endregion

        #region fsSelection fallback — only when the name table has nothing usable

        [TestMethod]
        public void GetSubfamilyEnum_NoSubfamilyRecords_FallsBackToFsSelection()
        {
            const ushort bold = 0x0020;
            const ushort italic = 0x0001;

            Assert.AreEqual(FontSubFamily.Regular, EnumForFsSelection(0));
            Assert.AreEqual(FontSubFamily.Bold, EnumForFsSelection(bold));
            Assert.AreEqual(FontSubFamily.Italic, EnumForFsSelection(italic));
            Assert.AreEqual(FontSubFamily.BoldItalic, EnumForFsSelection((ushort)(bold | italic)));
        }

        #endregion

        #region Helpers

        private static FontSubFamily EnumFor(string subfamilyName)
        {
            return CreateNameTable(
                MakeEnglishRecord(NameRecordTypes.FontSubfamilyName, subfamilyName))
                .GetSubfamilyEnum();
        }

        private static FontSubFamily EnumForFsSelection(ushort fsSelection)
        {
            var nameTable = CreateNameTable();
            nameTable.Os2FsSelection = fsSelection;
            return nameTable.GetSubfamilyEnum();
        }

        private static NameTable CreateNameTable(params NameRecord[] records)
        {
            return new NameTable { NameRecords = records };
        }

        /// <summary>
        /// Builds a Windows/en-US name record. GetEnglishName() matches on LanguageMapping,
        /// not on the raw languageID, so LanguageMapping must be populated for the
        /// English-preference tests to mean anything.
        /// </summary>
        private static NameRecord MakeEnglishRecord(NameRecordTypes type, string name)
        {
            const int enUs = 0x0409;
            return new NameRecord
            {
                RecordType = type,
                nameId = (ushort)type,
                platformId = 3,          // Windows
                encodingId = 1,          // Unicode BMP
                languageID = enUs,
                Name = name,
                LanguageMapping = new LanguageMapping { code = enUs, Language = Languages.English }
            };
        }

        /// <summary>
        /// Builds a non-English name record. The specific language is irrelevant to the logic
        /// under test — all that matters is that it is not Languages.English.
        /// </summary>
        private static NameRecord MakeLocalizedRecord(NameRecordTypes type, string name)
        {
            const int svSe = 0x041D;
            return new NameRecord
            {
                RecordType = type,
                nameId = (ushort)type,
                platformId = 3,
                encodingId = 1,
                languageID = svSe,
                Name = name,
                LanguageMapping = new LanguageMapping { code = svSe, Language = Languages.Swedish }
            };
        }

        #endregion
    }
}