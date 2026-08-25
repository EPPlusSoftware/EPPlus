/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/25/2026         EPPlus Software AB           Initial tests for NameTable subfamily mapping
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Name;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.FontScanning
{
    /// <summary>
    /// Low-level unit tests for <see cref="NameTable.GetSubfamilyEnum"/> — bare table instances,
    /// no font file loading. These tests exercise the string-to-enum heuristic directly, so they
    /// stay fast and deterministic regardless of which fonts are installed on the test machine.
    ///
    /// Background: fonts whose subfamily name carries a weight beyond Bold (e.g. "Black", "Heavy",
    /// "Demi" — as seen in real-world builds of "Arial Black") were previously mapped to
    /// FontSubFamily.Bold. That made an exact match against a Regular request impossible, so a
    /// query for "Arial Black" + Regular fell through to the fallback chain even though the font
    /// was installed and found by FindBestMatch. See NameTable.GetSubfamilyEnum for the fix.
    /// </summary>
    [TestClass]
    public class NameTableSubfamilyTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Weight names beyond Bold — should map to Regular, not Bold

        [TestMethod]
        public void GetSubfamilyEnum_Black_ReturnsRegular()
        {
            var nameTable = CreateNameTable(subfamilyName: "Black");

            var result = nameTable.GetSubfamilyEnum();

            Assert.AreEqual(FontSubFamily.Regular, result,
                "'Black' is a separate typographic weight, not a Bold variant. The family " +
                "name (e.g. 'Arial Black') already distinguishes it, so within the 4-value " +
                "FontSubFamily enum it must be treated as Regular.");
        }

        [TestMethod]
        public void GetSubfamilyEnum_Heavy_ReturnsRegular()
        {
            var nameTable = CreateNameTable(subfamilyName: "Heavy");

            var result = nameTable.GetSubfamilyEnum();

            Assert.AreEqual(FontSubFamily.Regular, result);
        }

        [TestMethod]
        public void GetSubfamilyEnum_Demi_ReturnsRegular()
        {
            var nameTable = CreateNameTable(subfamilyName: "Demi");

            var result = nameTable.GetSubfamilyEnum();

            Assert.AreEqual(FontSubFamily.Regular, result);
        }

        [TestMethod]
        public void GetSubfamilyEnum_Light_ReturnsRegular()
        {
            var nameTable = CreateNameTable(subfamilyName: "Light");

            var result = nameTable.GetSubfamilyEnum();

            Assert.AreEqual(FontSubFamily.Regular, result);
        }

        #endregion

        #region True RIBBI styles — must keep working (regression guard)

        [TestMethod]
        public void GetSubfamilyEnum_Regular_ReturnsRegular()
        {
            var nameTable = CreateNameTable(subfamilyName: "Regular");

            Assert.AreEqual(FontSubFamily.Regular, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetSubfamilyEnum_Bold_ReturnsBold()
        {
            var nameTable = CreateNameTable(subfamilyName: "Bold");

            Assert.AreEqual(FontSubFamily.Bold, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetSubfamilyEnum_Italic_ReturnsItalic()
        {
            var nameTable = CreateNameTable(subfamilyName: "Italic");

            Assert.AreEqual(FontSubFamily.Italic, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetSubfamilyEnum_BoldItalic_ReturnsBoldItalic()
        {
            var nameTable = CreateNameTable(subfamilyName: "Bold Italic");

            Assert.AreEqual(FontSubFamily.BoldItalic, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetSubfamilyEnum_SemiBold_StillContainsBold_ReturnsBold()
        {
            // "Semibold" legitimately contains the substring "bold" and is a reasonable
            // approximation of Bold within the 4-value enum — unlike "Black"/"Heavy"/"Demi",
            // which don't contain "bold" at all. This must keep matching Bold.
            var nameTable = CreateNameTable(subfamilyName: "SemiBold");

            Assert.AreEqual(FontSubFamily.Bold, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetSubfamilyEnum_Oblique_ReturnsItalic()
        {
            var nameTable = CreateNameTable(subfamilyName: "Oblique");

            Assert.AreEqual(FontSubFamily.Italic, nameTable.GetSubfamilyEnum());
        }

        #endregion

        #region Field priority — Typographic Subfamily (17) over legacy Subfamily (2)

        [TestMethod]
        public void GetSubfamilyEnum_PrefersTypographicSubfamily17OverLegacySubfamily2()
        {
            // Real-world "Arial Black"-style layout: legacy subfamily (2) carries the extra
            // weight name, while the typographic subfamily (17) correctly says "Regular".
            // GetSubfamilyName() must prefer 17, so the enum should resolve to Regular even
            // without the Black/Heavy/Demi fix above.
            var nameTable = new NameTable
            {
                NameRecords = new[]
                {
                    MakeRecord(NameRecordTypes.FontSubfamilyName, "Black"),
                    MakeRecord(NameRecordTypes.TypographicSubfamilyName, "Regular"),
                }
            };

            Assert.AreEqual(FontSubFamily.Regular, nameTable.GetSubfamilyEnum());
        }

        #endregion

        #region Fallback to OS/2 fsSelection when name table has no usable subfamily

        [TestMethod]
        public void GetSubfamilyEnum_NoNameRecords_FallsBackToFsSelectionBold()
        {
            const ushort fsSelectionBold = 0x0020;
            var nameTable = new NameTable
            {
                NameRecords = new NameRecord[0],
                Os2FsSelection = fsSelectionBold
            };

            Assert.AreEqual(FontSubFamily.Bold, nameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void GetSubfamilyEnum_NoNameRecords_FallsBackToFsSelectionRegular()
        {
            var nameTable = new NameTable
            {
                NameRecords = new NameRecord[0],
                Os2FsSelection = 0
            };

            Assert.AreEqual(FontSubFamily.Regular, nameTable.GetSubfamilyEnum());
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Builds a minimal bare NameTable with a single Font Subfamily Name (nameID 2) record —
        /// enough to exercise GetSubfamilyEnum's string-matching heuristic in isolation.
        /// </summary>
        private static NameTable CreateNameTable(string subfamilyName)
        {
            return new NameTable
            {
                NameRecords = new[]
                {
                    MakeRecord(NameRecordTypes.FontSubfamilyName, subfamilyName)
                }
            };
        }

        private static NameRecord MakeRecord(NameRecordTypes type, string name)
        {
            return new NameRecord
            {
                RecordType = type,
                nameId = (ushort)type,
                platformId = 3,
                encodingId = 1,
                Name = name
            };
        }

        #endregion
    }
}