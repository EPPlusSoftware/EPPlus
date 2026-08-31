/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/25/2026         EPPlus Software AB           Integration test for Arial Black scanning fix
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.FontScanning
{
    /// <summary>
    /// Integration-level tests for the Arial Black scanning bug: a font whose subfamily name
    /// is "Black" was mapped to FontSubFamily.Bold, which made an exact match against a
    /// Regular request impossible even though the font was installed and found on disk.
    ///
    /// Unlike NameTableSubfamilyTests (bare table, no I/O), these tests go through the full
    /// scanning pipeline (FontScannerV2.FindBestMatch / OpenTypeFontEngine.GetFontAvailability)
    /// against a real, installed "Arial Black" font file — so they depend on that font actually
    /// being present on the machine running the tests. Uses RequireFont / SystemFontsEngine
    /// (see FontTestBase) to mark itself Inconclusive rather than fail on machines that lack it,
    /// consistent with the other system-font-dependent tests in this suite (e.g. Goudy Stout,
    /// Aptos Narrow).
    /// </summary>
    [TestClass]
    public class ArialBlackScanningIntegrationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void FindBestMatch_ArialBlackRegular_ScansSubfamilyAsRegular()
        {
            RequireFont(SystemFontsEngine, "Arial Black", FontSubFamily.Regular);

            var face = FontScannerV2.FindBestMatch(
                FontFolders, "Arial Black", FontSubFamily.Regular, searchSystemDirectories: true);

            Assert.IsNotNull(face, "Arial Black should have been found on disk.");
            Assert.AreEqual("Arial Black", face.FamilyName);
            Assert.AreEqual(FontSubFamily.Regular, face.Subfamily,
                "Arial Black's base instance must scan as Regular, not Bold — its 'Black' " +
                "weight is already captured by FamilyName, not by FontSubFamily.");
        }

        [TestMethod]
        public void FindBestMatch_ArialBlackRegular_IsExactMatch()
        {
            RequireFont(SystemFontsEngine, "Arial Black", FontSubFamily.Regular);

            var face = FontScannerV2.FindBestMatch(
                FontFolders, "Arial Black", FontSubFamily.Regular, searchSystemDirectories: true);

            Assert.IsNotNull(face);
            Assert.IsTrue(face.IsExactMatch,
                "A request for 'Arial Black' + Regular against the installed 'Arial Black' " +
                "font must score as an exact match, or callers relying on IsExactMatch " +
                "(e.g. DefaultFontResolver) will incorrectly fall through to the fallback chain.");
        }

        [TestMethod]
        public void GetFontAvailability_ArialBlackRegular_ReturnsExact()
        {
            RequireFont(SystemFontsEngine, "Arial Black", FontSubFamily.Regular);

            // End-to-end: this is the same call DefaultFontResolver.ResolveFont relies on
            // before falling through to the user/built-in fallback chain and, ultimately,
            // Archivo Narrow. Before the fix this returned FamilyOnly (or NotFound), which
            // is exactly what sent "Arial Black" through to Liberation Sans in production.
            var availability = SystemFontsEngine.GetFontAvailability("Arial Black", FontSubFamily.Regular);

            Assert.AreEqual(FontAvailability.Exact, availability);
        }
    }
}