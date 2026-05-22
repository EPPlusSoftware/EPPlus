/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2026         EPPlus Software AB           Boundary tests for binary-search classifier
 *************************************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests
{
    /// <summary>
    /// Boundary tests for <see cref="UnicodeScriptClassifier"/>.
    ///
    /// We do not test every script's content — that would be source-code duplication, since
    /// the classifier is effectively a static lookup table. Instead we verify that the binary
    /// search treats range endpoints as inclusive on both sides, which protects against
    /// off-by-one bugs if the algorithm is rewritten.
    ///
    /// Functional correctness of the table is exercised indirectly by higher-level tests
    /// (DefaultFontProvider routing emoji to Noto Emoji, etc.).
    /// </summary>
    [TestClass]
    public class UnicodeScriptClassifierTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void Classify_CodePointOnRangeStart_IsInclusive()
        {
            // U+4E00 is the first code point of the CJK Unified Ideographs main block.
            // If the binary search treats Start as exclusive, this returns Unknown.
            Assert.AreEqual(UnicodeScript.Han, UnicodeScriptClassifier.OfCodePoint(0x4E00));
        }

        [TestMethod]
        public void Classify_CodePointOnRangeEnd_IsInclusive()
        {
            // U+9FFF is the last code point of the CJK Unified Ideographs main block.
            // If the binary search treats End as exclusive, this returns Unknown.
            Assert.AreEqual(UnicodeScript.Han, UnicodeScriptClassifier.OfCodePoint(0x9FFF));
        }

        [TestMethod]
        public void Classify_CodePointJustOutsideRange_ReturnsUnknown()
        {
            // U+036F is the last Combining Diacritical Mark; U+0370 starts Greek.
            // Verifies that the classifier does not over-shoot a range on the high side.
            Assert.AreEqual(UnicodeScript.Unknown, UnicodeScriptClassifier.OfCodePoint(0x036F));
        }
    }
}