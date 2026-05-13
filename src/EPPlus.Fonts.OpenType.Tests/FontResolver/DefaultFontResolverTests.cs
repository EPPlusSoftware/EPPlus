/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/06/2026         EPPlus Software AB           DefaultFontResolver unit tests
  05/06/2026         EPPlus Software AB           Updated for property-based EpplusFontConfiguration
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.FontResolver
{
    /// <summary>
    /// Unit tests for DefaultFontResolver, covering the four-step resolution flow:
    ///   1. Exact match
    ///   2. User-configured fallback chain
    ///   3. Built-in fallback chain
    ///   4. Archivo Narrow (embedded ultimate fallback)
    ///
    /// Tests use FakeFontScanner and FakeFontFileReader to control exactly which fonts
    /// "exist" on the system, making behavior fully deterministic and platform-independent.
    /// </summary>
    [TestClass]
    public class DefaultFontResolverTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Step 1 — Exact match

        [TestMethod]
        public void ResolveFont_ExactMatchExists_ReturnsExactMatch()
        {
            // Arrange
            var scanner = new FakeFontScanner()
                .Register("Calibri", FontSubFamily.Regular, "fake://Calibri.ttf");
            var reader = new FakeFontFileReader()
                .Register("fake://Calibri.ttf", "FAKE:Calibri");

            var resolver = new DefaultFontResolver(scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Regular);

            // Assert
            AssertMarker("FAKE:Calibri", bytes);
        }

        #endregion

        #region Step 3 — Built-in fallback chain

        [TestMethod]
        public void ResolveFont_NoExactMatch_FallsThroughToBuiltinChain()
        {
            // Arrange — Calibri itself is unavailable, but Carlito (first in built-in chain) is.
            var scanner = new FakeFontScanner()
                .Register("Carlito", FontSubFamily.Regular, "fake://Carlito.ttf");
            var reader = new FakeFontFileReader()
                .Register("fake://Carlito.ttf", "FAKE:Carlito");

            var resolver = new DefaultFontResolver(scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Regular);

            // Assert
            AssertMarker("FAKE:Carlito", bytes);
        }

        [TestMethod]
        public void ResolveFont_BuiltinChainSkipsMissing_TriesNextEntry()
        {
            // Arrange — Calibri chain is Carlito → Liberation Sans → Arial → Helvetica.
            // Skip the first two and verify Arial is selected.
            var scanner = new FakeFontScanner()
                .Register("Arial", FontSubFamily.Regular, "fake://Arial.ttf");
            var reader = new FakeFontFileReader()
                .Register("fake://Arial.ttf", "FAKE:Arial");

            var resolver = new DefaultFontResolver(scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Regular);

            // Assert
            AssertMarker("FAKE:Arial", bytes);
        }

        [TestMethod]
        public void ResolveFont_BoldRequest_OnlyMatchesBoldFaces()
        {
            // Arrange — request Calibri Bold. Carlito Regular is available but Carlito Bold
            // is not. The resolver must NOT downgrade to Carlito Regular; it should walk
            // further down the chain.
            var scanner = new FakeFontScanner()
                .Register("Carlito", FontSubFamily.Regular, "fake://Carlito-Regular.ttf")
                .Register("Arial", FontSubFamily.Bold, "fake://Arial-Bold.ttf");
            var reader = new FakeFontFileReader()
                .Register("fake://Carlito-Regular.ttf", "FAKE:CarlitoRegular")
                .Register("fake://Arial-Bold.ttf", "FAKE:ArialBold");

            var resolver = new DefaultFontResolver(scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Bold);

            // Assert — must reach Arial Bold, not pick up Carlito Regular along the way
            AssertMarker("FAKE:ArialBold", bytes);
        }

        #endregion

        #region Step 2 — User-configured fallback chain (precedence)

        [TestMethod]
        public void ResolveFont_UserConfigBeforeBuiltin_UserConfigWins()
        {
            // Arrange — user says "if Calibri is missing, use MyCustomFont". MyCustomFont exists.
            // The built-in chain would otherwise pick Carlito (which also exists). Verify the
            // user's choice wins.
            var config = new EpplusFontConfiguration();
            config.FontFallbacks["Calibri"] = new[] { "MyCustomFont" };

            var scanner = new FakeFontScanner()
                .Register("MyCustomFont", FontSubFamily.Regular, "fake://MyCustomFont.ttf")
                .Register("Carlito", FontSubFamily.Regular, "fake://Carlito.ttf");
            var reader = new FakeFontFileReader()
                .Register("fake://MyCustomFont.ttf", "FAKE:MyCustomFont")
                .Register("fake://Carlito.ttf", "FAKE:Carlito");

            var resolver = new DefaultFontResolver(config: config, scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Regular);

            // Assert
            AssertMarker("FAKE:MyCustomFont", bytes);
        }

        [TestMethod]
        public void ResolveFont_UserConfigMisses_BuiltinTakesOver()
        {
            // Arrange — user has configured Calibri → NonExistent, but NonExistent isn't on the
            // system. Built-in chain Carlito → ... should kick in and find Carlito.
            var config = new EpplusFontConfiguration();
            config.FontFallbacks["Calibri"] = new[] { "NonExistent" };

            var scanner = new FakeFontScanner()
                .Register("Carlito", FontSubFamily.Regular, "fake://Carlito.ttf");
            var reader = new FakeFontFileReader()
                .Register("fake://Carlito.ttf", "FAKE:Carlito");

            var resolver = new DefaultFontResolver(config: config, scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Regular);

            // Assert
            AssertMarker("FAKE:Carlito", bytes);
        }

        #endregion

        #region Step 4 — Archivo Narrow fallback

        [TestMethod]
        public void ResolveFont_NoMatchAnywhere_FallsBackToArchivoNarrow()
        {
            // Arrange — nothing exists on the system, no user config, no built-in chain entry
            // for the requested font (it's a made-up name).
            var scanner = new FakeFontScanner();
            var reader = new FakeFontFileReader();

            var resolver = new DefaultFontResolver(scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("DefinitelyNotARealFont_XYZ", FontSubFamily.Regular);

            // Assert — bytes must be a real font (Archivo Narrow), not a fake marker.
            // Validate by checking that the bytes parse as a real OpenType font with the
            // expected family name.
            Assert.IsNotNull(bytes);
            Assert.IsTrue(bytes.Length > 1000, "Archivo Narrow should be a real font file, much larger than any fake marker");

            var parsedFont = TestFolderEngine.GetFromBytes(bytes);
            Assert.AreEqual("Archivo Narrow", parsedFont.NameTable.GetFamilyName());
            Assert.AreEqual(FontSubFamily.Regular, parsedFont.NameTable.GetSubfamilyEnum());
        }

        [TestMethod]
        public void ResolveFont_BuiltinChainExistsButNothingInstalled_FallsBackToArchivoNarrow()
        {
            // Arrange — Calibri has a built-in chain but NONE of the chain entries exist on
            // the system. Should fall through to Archivo Narrow as ultimate safety net.
            var scanner = new FakeFontScanner();
            var reader = new FakeFontFileReader();

            var resolver = new DefaultFontResolver(scanner: scanner, fileReader: reader);

            // Act
            var bytes = resolver.ResolveFont("Calibri", FontSubFamily.Regular);

            // Assert
            var parsedFont = TestFolderEngine.GetFromBytes(bytes);
            Assert.AreEqual("Archivo Narrow", parsedFont.NameTable.GetFamilyName());
        }

        #endregion

        #region Helpers

        /// <summary>
        /// Asserts that the returned bytes match the UTF-8 encoding of the expected marker.
        /// FakeFontFileReader.Register(path, string) writes UTF-8 bytes for the marker, so
        /// this is the inverse check.
        /// </summary>
        private static void AssertMarker(string expectedMarker, byte[] actualBytes)
        {
            Assert.IsNotNull(actualBytes, "ResolveFont returned null");
            var expected = Encoding.UTF8.GetBytes(expectedMarker);
            CollectionAssert.AreEqual(expected, actualBytes,
                "Expected marker '" + expectedMarker + "' but got " + Encoding.UTF8.GetString(actualBytes));
        }

        #endregion
    }
}