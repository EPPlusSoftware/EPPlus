/*************************************************************************************************
  Font Provider Unit Tests
  Tests for automatic emoji fallback functionality and script-based glyph fallback.
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.FallbackFonts
{
    [TestClass]
    public class FontProviderTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private OpenTypeFont _robotoFont;

        // U+6F22 = 漢 (a common Han ideograph used in Chinese/Japanese)
        private const string HanCharacter = "\u6F22";

        [TestInitialize]
        public void TestSetup()
        {
            _robotoFont = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular);
        }

        [TestMethod]
        public void DefaultFontProvider_EmojiGlyph_ShouldUseFallbackFont()
        {
            // Arrange
            var shaper = new TextShaper(TestFolderEngine, _robotoFont);

            // Act
            var shaped = shaper.Shape("😀");
            var usedFonts = shaper.GetUsedFonts().ToList();

            // Assert
            Assert.AreEqual(1, shaped.Glyphs.Length, "Should have 1 glyph");
            Assert.AreNotEqual((ushort)0, shaped.Glyphs[0].GlyphId, "Emoji should not be .notdef");
            Assert.AreNotEqual((byte)0, shaped.Glyphs[0].FontId, "Emoji should come from a fallback font, not primary");

            // Primary is always registered at FontId 0, even if no glyphs come from it.
            // The emoji fallback occupies FontId 1.
            Assert.AreEqual(2, usedFonts.Count, "Used fonts should be [primary, emoji fallback]");
            Assert.AreEqual(_robotoFont, usedFonts[0], "Primary font is always FontId 0");
            Assert.AreNotEqual(_robotoFont, usedFonts[1], "Fallback should be the emoji font");
        }

        [TestMethod]
        public void DefaultFontProvider_LatinText_ShouldUsePrimaryFont()
        {
            // Arrange
            var shaper = new TextShaper(TestFolderEngine, _robotoFont);

            // Act
            var shaped = shaper.Shape("Hello World");

            // Assert
            foreach (var glyph in shaped.Glyphs)
            {
                Assert.AreEqual((byte)0, glyph.FontId, "All glyphs should be from primary font");
            }
        }

        [TestMethod]
        public void DefaultFontProvider_MixedTextAndEmoji_ShouldUseMultipleFonts()
        {
            // Arrange
            var shaper = new TextShaper(TestFolderEngine, _robotoFont);

            // Act
            var shaped = shaper.Shape("Hello 😀 World");
            var usedFonts = shaper.GetUsedFonts().ToList();

            // Assert
            Assert.AreEqual(2, usedFonts.Count, "Should use 2 fonts (primary + emoji fallback)");
            Assert.AreEqual(_robotoFont, usedFonts[0], "First font should be primary");
            Assert.AreNotEqual(_robotoFont, usedFonts[1], "Second font should be emoji fallback");
        }

        [TestMethod]
        public void TextShaper_SurrogatePair_ShouldMapToSingleGlyph()
        {
            // Arrange
            var shaper = new TextShaper(TestFolderEngine, _robotoFont);
            string text = "😀"; // U+1F600 = 2 chars in UTF-16

            // Act
            var shaped = shaper.Shape(text);

            // Assert
            Assert.AreEqual(2, text.Length, "Emoji should be 2 chars in .NET");
            Assert.AreEqual(1, shaped.Glyphs.Length, "Should map to 1 glyph");
            Assert.AreEqual((byte)2, shaped.Glyphs[0].CharCount, "Glyph should span 2 chars");
            Assert.AreEqual((ushort)0, shaped.Glyphs[0].ClusterIndex, "Should start at char 0");
        }

        [TestMethod]
        public void TextShaper_MultipleEmoji_ShouldMapCorrectly()
        {
            // Arrange
            var shaper = new TextShaper(TestFolderEngine, _robotoFont);
            string text = "😀😁😂"; // 3 emoji = 6 chars in UTF-16

            // Act
            var shaped = shaper.Shape(text);

            // Assert
            Assert.AreEqual(6, text.Length, "3 emoji = 6 chars");
            Assert.AreEqual(3, shaped.Glyphs.Length, "Should map to 3 glyphs");

            Assert.AreEqual((ushort)0, shaped.Glyphs[0].ClusterIndex, "First emoji at char 0");
            Assert.AreEqual((byte)2, shaped.Glyphs[0].CharCount, "First emoji spans 2 chars");

            Assert.AreEqual((ushort)2, shaped.Glyphs[1].ClusterIndex, "Second emoji at char 2");
            Assert.AreEqual((byte)2, shaped.Glyphs[1].CharCount, "Second emoji spans 2 chars");

            Assert.AreEqual((ushort)4, shaped.Glyphs[2].ClusterIndex, "Third emoji at char 4");
            Assert.AreEqual((byte)2, shaped.Glyphs[2].CharCount, "Third emoji spans 2 chars");
        }

        // -----------------------------------------------------------------------------------------
        // Script-based glyph fallback
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void DefaultFontProvider_HanGlyphWithConfiguredFallback_RoutesToFallbackFont()
        {
            // Arrange — explicit script fallback to BIZ UDGothic (which is in the test font folder
            // and contains Han glyphs). Use a fresh engine so the default Han chain doesn't
            // interfere.
            var engine = new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.SetScriptFallback(UnicodeScript.Han, "BIZ UDGothic");
            });

            var roboto = engine.LoadFont("Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(engine, roboto);

            // Act
            var shaped = shaper.Shape(HanCharacter);
            var usedFonts = shaper.GetUsedFonts().ToList();

            // Assert — glyph must come from the fallback, not from Roboto's .notdef
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreNotEqual((ushort)0, shaped.Glyphs[0].GlyphId, "Han glyph should not be .notdef");
            Assert.AreNotEqual((byte)0, shaped.Glyphs[0].FontId, "Han glyph should come from a fallback font");

            Assert.AreEqual(2, usedFonts.Count, "Should use 2 fonts (primary + Han fallback)");
            Assert.AreEqual(roboto, usedFonts[0]);
            Assert.AreEqual("BIZ UDGothic", usedFonts[1].NameTable.GetFamilyName());
        }

        [TestMethod]
        public void DiagnoseScriptFallback()
        {
            var engine = new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.SetScriptFallback(UnicodeScript.Han, "BIZ UDGothic");
            });

            // Diagnostik 1: vad sätter engine för Han efter konstruktion?
            var chain = engine.GetScriptFallback(UnicodeScript.Han);
            System.Console.WriteLine($"[DIAG] Han chain: [{string.Join(", ", chain ?? new string[0])}]");

            // Diagnostik 2: shape och se vad providern returnerar
            var roboto = engine.LoadFont("Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(engine, roboto);
            var shaped = shaper.Shape("\u6F22");

            System.Console.WriteLine($"[DIAG] Glyph count: {shaped.Glyphs.Length}");
            System.Console.WriteLine($"[DIAG] First glyph: GlyphId={shaped.Glyphs[0].GlyphId}, FontId={shaped.Glyphs[0].FontId}");

            var usedFonts = shaper.GetUsedFonts().ToList();
            System.Console.WriteLine($"[DIAG] Used fonts count: {usedFonts.Count}");
            foreach (var f in usedFonts)
            {
                System.Console.WriteLine($"[DIAG] - {f.NameTable.GetFamilyName()}");
            }
        }

        [TestMethod]
        public void DefaultFontProvider_HanGlyphWithNoInstalledFallback_ReturnsNotdef()
        {
            // Arrange — TestFolderEngine's default Han chain points at Microsoft YaHei, SimSun, etc.,
            // none of which are present in the test font folder. So a Han character should fall
            // through to .notdef.
            var roboto = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(TestFolderEngine, roboto);

            // Act
            var shaped = shaper.Shape(HanCharacter);

            // Assert
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreEqual((ushort)0, shaped.Glyphs[0].GlyphId, "Glyph should be .notdef");
            Assert.AreEqual((byte)0, shaped.Glyphs[0].FontId, "Provider returns primary font when not found");
        }

        [TestMethod]
        public void DefaultFontProvider_EmptyScriptFallbackChain_DisablesFallbackForScript()
        {
            // Arrange — explicitly disable Han fallback. Even if BIZ UDGothic would have worked
            // by default, an empty chain says "do not route Han characters anywhere".
            var engine = new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.SetScriptFallback(UnicodeScript.Han, new string[0]);
            });

            var roboto = engine.LoadFont("Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(engine, roboto);

            // Act
            var shaped = shaper.Shape(HanCharacter);

            // Assert
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreEqual((ushort)0, shaped.Glyphs[0].GlyphId, "Glyph should be .notdef when chain is empty");
        }

        [TestMethod]
        public void EpplusFontConfiguration_Reset_RestoresDefaultScriptChains()
        {
            // Arrange — start with a custom Han chain, then call Reset on the configuration.
            // After Reset, the default Han chain (Microsoft YaHei, SimSun, Noto Sans CJK SC,
            // PingFang SC) should be reinstated.
            //
            // We can't directly observe the chain from outside the engine, so we verify the
            // behavior indirectly: with the test font folder only and the default chain back
            // in place, a Han character should once again be .notdef (none of the default
            // chain fonts are in the test folder).
            var engine = new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.SetScriptFallback(UnicodeScript.Han, "BIZ UDGothic");

                // Custom chain was just set above — now undo via Reset.
                cfg.Reset();

                // After Reset we lose the font directories too, so re-add them.
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
            });

            var roboto = engine.LoadFont("Roboto", FontSubFamily.Regular);
            var shaper = new TextShaper(engine, roboto);

            // Act
            var shaped = shaper.Shape(HanCharacter);

            // Assert — back to default chain. None of (Microsoft YaHei, SimSun, ...) is in the
            // test font folder, so the Han character lands on .notdef.
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreEqual((ushort)0, shaped.Glyphs[0].GlyphId,
                "After Reset, the default Han chain should be active again; with no matching fonts installed, glyph is .notdef");
        }

        [TestMethod]
        public void DefaultFontProvider_ChineseTextWithPunctuationInAptosNarrow_AllGlyphsResolve()
        {
            // Reproduces a real-world bug found during PDF export: Chinese text with CJK
            // punctuation rendered in Aptos Narrow shows .notdef glyphs for the punctuation
            // marks. The Han ideographs route to the Han fallback chain correctly, but the
            // fullwidth comma (U+FF0C) and ideographic full stop (U+3002) are classified as
            // Unknown by UnicodeScriptClassifier and never reach any fallback font.
            //
            // After the classifier is fixed to include U+3000-U+303F and U+FF00-U+FFEF in
            // the Han range, all 22 glyphs should resolve to non-zero glyph ids.

            // Arrange
            RequireFont(SystemFontsEngine, "Aptos Narrow", FontSubFamily.Regular);

            var primary = SystemFontsEngine.LoadFont("Aptos Narrow", FontSubFamily.Regular);
            var shaper = new TextShaper(SystemFontsEngine, primary);

            string text = "我今天吃了太多饺子，现在我看起来像一个饺子。";

            // Act
            var shaped = shaper.Shape(text);

            // Assert — every glyph in the shaped output must be a real glyph, not .notdef.
            // Identify which (if any) failed so the diagnostic message points at the actual
            // characters that landed on .notdef.
            var notdefs = new System.Collections.Generic.List<string>();
            for (int i = 0; i < shaped.Glyphs.Length; i++)
            {
                if (shaped.Glyphs[i].GlyphId == 0)
                {
                    int charIndex = shaped.Glyphs[i].ClusterIndex;
                    char ch = text[charIndex];
                    notdefs.Add(string.Format("index {0}: U+{1:X4} '{2}'", charIndex, (int)ch, ch));
                }
            }

            Assert.AreEqual(0, notdefs.Count,
                "All glyphs should resolve to real glyph ids. The following characters landed on .notdef: "
                + string.Join(", ", notdefs));
        }

        // -----------------------------------------------------------------------------------------
        // CJK punctuation routing — reproduces a real-world bug found during PDF export of
        // Chinese text rendered with Aptos Narrow. Fullwidth/CJK punctuation should route to the
        // Han fallback chain, but currently lands on .notdef because these code-point ranges are
        // missing from UnicodeScriptClassifier.
        //
        // These tests verify the bug first; the classifier fix follows in a separate change.
        // -----------------------------------------------------------------------------------------

        // U+6211 = 我 (Han ideograph)
        private const string HanIdeograph = "\u6211";

        // U+FF0C = ， (Fullwidth Comma — used as the CJK comma in Chinese / Japanese / Korean)
        private const string CjkFullwidthComma = "\uFF0C";

        // U+3002 = 。 (Ideographic Full Stop — used as the CJK period)
        private const string CjkIdeographicFullStop = "\u3002";

        [TestMethod]
        public void DefaultFontProvider_HanIdeographInAptosNarrow_RoutesToHanFallback()
        {
            // Arrange — use the system engine because Aptos Narrow is a system font (not in the
            // test font folder). Skip the test if Aptos Narrow is not available on this machine.
            RequireFont(SystemFontsEngine, "Aptos Narrow", FontSubFamily.Regular);

            var primary = SystemFontsEngine.LoadFont("Aptos Narrow", FontSubFamily.Regular);
            var shaper = new TextShaper(SystemFontsEngine, primary);

            // Act
            var shaped = shaper.Shape(HanIdeograph);

            // Assert — sanity check: a Han ideograph should route to the Han fallback chain.
            // If this fails, the entire script routing for Han is broken, not just punctuation.
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreNotEqual((ushort)0, shaped.Glyphs[0].GlyphId,
                "Han ideograph should not be .notdef — it should route to the Han fallback chain");
        }

        [TestMethod]
        public void DefaultFontProvider_CjkFullwidthCommaInAptosNarrow_RoutesToHanFallback()
        {
            // Arrange
            RequireFont(SystemFontsEngine, "Aptos Narrow", FontSubFamily.Regular);

            var primary = SystemFontsEngine.LoadFont("Aptos Narrow", FontSubFamily.Regular);
            var shaper = new TextShaper(SystemFontsEngine, primary);

            // Act
            var shaped = shaper.Shape(CjkFullwidthComma);

            // Assert — the fullwidth comma is shared CJK punctuation. It should route to the
            // Han fallback chain along with the rest of the CJK text it accompanies. Currently
            // U+FF00–U+FFEF is missing from UnicodeScriptClassifier, so the comma is classified
            // as Unknown and lands on .notdef.
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreNotEqual((ushort)0, shaped.Glyphs[0].GlyphId,
                "CJK fullwidth comma (U+FF0C) should not be .notdef — it should route to the Han fallback chain");
        }

        [TestMethod]
        public void DefaultFontProvider_CjkIdeographicFullStopInAptosNarrow_RoutesToHanFallback()
        {
            // Arrange
            RequireFont(SystemFontsEngine, "Aptos Narrow", FontSubFamily.Regular);

            var primary = SystemFontsEngine.LoadFont("Aptos Narrow", FontSubFamily.Regular);
            var shaper = new TextShaper(SystemFontsEngine, primary);

            // Act
            var shaped = shaper.Shape(CjkIdeographicFullStop);

            // Assert — the ideographic full stop is shared CJK punctuation. Currently
            // U+3000–U+303F is missing from UnicodeScriptClassifier, so the period is
            // classified as Unknown and lands on .notdef.
            Assert.AreEqual(1, shaped.Glyphs.Length);
            Assert.AreNotEqual((ushort)0, shaped.Glyphs[0].GlyphId,
                "CJK ideographic full stop (U+3002) should not be .notdef — it should route to the Han fallback chain");
        }
    }
}