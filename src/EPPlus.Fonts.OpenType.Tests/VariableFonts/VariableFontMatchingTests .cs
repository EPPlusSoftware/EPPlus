/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/25/2026         EPPlus Software AB           Variable font matching tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.VariableFonts
{
    /// <summary>
    /// Verifies that the font scanner treats variable fonts as capable of delivering only
    /// their default named instance. A variable font must not masquerade as an exact match
    /// for a non-default subfamily.
    ///
    /// Regression context: a developer had Archivo Narrow installed as a wght-axis variable
    /// web font. The scanner returned that file for a Bold request, the font library then read
    /// the default (Regular) instance, and the developer wrote asserts against those Regular
    /// values. On machines without the variable font installed, resolution fell back to the
    /// embedded static Archivo Narrow Bold and the asserts failed — a non-deterministic,
    /// machine-dependent test failure.
    ///
    /// These tests are deliberately written against FontScannerV2 directly (not through an
    /// engine) so they exercise exactly the matching logic that changed, with no dependency on
    /// system-installed fonts and no interference from the Archivo Narrow special-case in
    /// DefaultFontResolver.
    /// </summary>
    [TestClass]
    public class VariableFontMatchingTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        // Variable (wght-axis) build of Archivo Narrow. Its default instance is Regular (wght 400).
        private const string VariableFontFamily = "Archivo Narrow";
        private const string VariableFontFileName = "ArchivoNarrow-VariableFont_wght.ttf";

        /// <summary>
        /// The isolated directory containing only the variable font. Pointing the scanner here
        /// (with system directories disabled) guarantees the variable face is the only candidate,
        /// which makes the disqualification observable as a null result.
        /// </summary>
        private static List<string> VariableFontDirectories
        {
            get { return new List<string> { Path.Combine(FontFolder, "VariableFonts") }; }
        }

        private static string VariableFontPath
        {
            get { return Path.Combine(FontFolder, "VariableFonts", VariableFontFileName); }
        }

        [TestMethod]
        public void ScanSingleFace_VariableFont_SetsIsVariable()
        {
            // The fvar table must be detected purely from the table directory.
            var face = FontScannerV2.GetFace(VariableFontPath);

            Assert.IsNotNull(face, "Expected the variable font to be scanned.");
            Assert.IsTrue(face.IsVariable,
                "A font containing an 'fvar' table must be flagged as variable.");
        }

        [TestMethod]
        public void FindBestMatch_VariableFont_DefaultStyle_IsExactMatch()
        {
            // The default instance of this variable font IS Regular, so a Regular request is a
            // legitimate exact match. This guards against over-penalising variable fonts: they
            // must still satisfy a request for their default subfamily.
            var match = FontScannerV2.FindBestMatch(
                VariableFontDirectories,
                VariableFontFamily,
                FontSubFamily.Regular,
                searchSystemDirectories: false);

            Assert.IsNotNull(match,
                "A variable font must still match a request for its default (Regular) subfamily.");
            Assert.IsTrue(match.IsVariable, "The matched face is expected to be variable.");
            Assert.AreEqual(FontSubFamily.Regular, match.Subfamily);
            Assert.IsTrue(match.IsExactMatch,
                "A variable font matching its default subfamily must be reported as an exact match.");
        }

        [TestMethod]
        public void FindBestMatch_VariableFont_NonDefaultStyle_IsDisqualified()
        {
            // Bold is NOT the default instance. Without variation interpolation the file cannot
            // deliver Bold, so the variable face must be disqualified. With no other candidate in
            // the isolated directory, the scanner returns null — and crucially never returns the
            // Regular face flagged as an exact Bold match (the original bug).
            var match = FontScannerV2.FindBestMatch(
                VariableFontDirectories,
                VariableFontFamily,
                FontSubFamily.Bold,
                searchSystemDirectories: false);

            Assert.IsNull(match,
                "A variable font whose default instance is not Bold must not be returned as a " +
                "match for a Bold request when it is the only candidate.");
        }

        [TestMethod]
        public void FindBestMatch_VariableFont_NonDefaultStyle_IsNotExactMatch()
        {
            // Belt-and-braces companion to the disqualification test, phrased as the property we
            // actually care about: even if some future change let a variable face survive as a
            // low-scoring candidate for a non-default style, it must never be flagged exact.
            var match = FontScannerV2.FindBestMatch(
                VariableFontDirectories,
                VariableFontFamily,
                FontSubFamily.BoldItalic,
                searchSystemDirectories: false);

            // Current behaviour: disqualified → null. If that ever changes, the match must at
            // least not be exact.
            if (match != null)
            {
                Assert.IsFalse(match.IsExactMatch,
                    "A variable font must never be an exact match for a non-default subfamily.");
            }
        }
    }
}