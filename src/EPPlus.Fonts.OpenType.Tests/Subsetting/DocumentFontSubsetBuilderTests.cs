using EPPlus.Fonts.OpenType.Subsetting;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class DocumentFontSubsetBuilderTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private static DocumentFontSubsetBuilder CreateBuilderWithCallback(
    Func<FontEmbeddingInfo, FontEmbeddingDecision> callback)
        {
            var engine = new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.OnFontEmbedding(callback);
            });
            return new DocumentFontSubsetBuilder(engine);
        }

        private static Func<FontEmbeddingInfo, FontEmbeddingDecision> SkipByName(string namePart)
        {
            return info => info.FontName != null && info.FontName.Contains(namePart)
                ? FontEmbeddingDecision.Skip
                : FontEmbeddingDecision.Default;
        }

        [TestMethod]
        public void Build_SkippedPrimary_NextFontBecomesPrimary()
        {
            var builder = CreateBuilderWithCallback(SkipByName("Roboto"));
            builder.AddText("Roboto", FontSubFamily.Regular, "Hello");
            builder.Build();

            var provider = builder.GetShapingProvider("Roboto", FontSubFamily.Regular);

            Assert.IsNotNull(provider.PrimaryFont);
            StringAssert.DoesNotMatch(
                provider.PrimaryFont.GetEnglishFontFamilyName(),
                new System.Text.RegularExpressions.Regex("Roboto"),
                "A skipped primary must not remain the provider's primary font.");
        }

        [TestMethod]
        public void Build_SkippedPrimary_AllTextSkipped_UsesLastResort()
        {
            var builder = CreateBuilderWithCallback(SkipByName("Roboto"));
            builder.AddText("Roboto", FontSubFamily.Regular, "Hello");
            builder.Build();

            var provider = builder.GetShapingProvider("Roboto", FontSubFamily.Regular);

            StringAssert.Contains(
                provider.PrimaryFont.GetEnglishFontFamilyName(),
                "Archivo",
                "When the whole chain is skipped, the last-resort font must become primary.");

            ushort glyphId;
            Assert.IsTrue(
                provider.PrimaryFont.CmapTable.TryGetGlyphId('H', out glyphId) && glyphId != 0,
                "Latin glyphs must be carried by the last-resort font after redistribution.");
        }

        [TestMethod]
        public void Build_SkippedPrimary_PrefersChainFontOverLastResort()
        {
            // Roboto skipped, but its text is an emoji that a chain fallback (Noto Emoji) covers.
            // The emoji must be carried by that chain font — NOT by the Archivo last resort.
            // (Han/CJK cannot be used here until script fallback is wired into the provider chain.)
            var builder = CreateBuilderWithCallback(SkipByName("Roboto"));
            builder.AddText("Roboto", FontSubFamily.Regular, char.ConvertFromUtf32(0x1F600)); // 😀
            builder.Build();

            var provider = builder.GetShapingProvider("Roboto", FontSubFamily.Regular);

            ushort glyphId;
            Assert.IsTrue(
                provider.PrimaryFont.CmapTable.TryGetGlyphId(0x1F600, out glyphId) && glyphId != 0,
                "The emoji must be carried by the chain fallback, not the last resort.");
        }

        [TestMethod]
        public void Build_SharedFallback_AllPrimariesSkipped_ProduceSingleConsistentSubset()
        {
            // The A1/B1/C1 regression, as a unit test: three primaries, all skipped, all collapsing
            // to the same last-resort font. That font must be ONE shared subset containing every
            // routed glyph — not three colliding subsets.
            var builder = CreateBuilderWithCallback(info => FontEmbeddingDecision.Skip); // skip everything
            builder.AddText("Roboto", FontSubFamily.Regular, "A");
            builder.AddText("Open Sans", FontSubFamily.Regular, "B");
            builder.AddText("Mulish", FontSubFamily.Regular, "C");
            builder.Build();

            var embedded = builder.GetFontsToEmbed().ToList();

            // Exactly one font embedded (the shared last resort), carrying A, B and C.
            Assert.AreEqual(1, embedded.Count, "All skipped primaries must collapse to one shared font.");
            var shared = embedded[0].Font;

            foreach (var ch in new[] { 'A', 'B', 'C' })
            {
                ushort glyphId;
                Assert.IsTrue(
                    shared.CmapTable.TryGetGlyphId(ch, out glyphId) && glyphId != 0,
                    "Shared subset must carry '" + ch + "' from all three primaries.");
            }
        }

        private const string TestFamily = "Roboto";

        [TestMethod]
        public void AddText_WithNullOrEmpty_DoesNotThrow()
        {
            var builder = new DocumentFontSubsetBuilder(TestFolderEngine);
            builder.AddText(TestFamily, FontSubFamily.Regular, null);
            builder.AddText(TestFamily, FontSubFamily.Regular, "");
            // No text was ever added, so Build has nothing to do — it must not throw either.
            builder.Build();
        }

        [TestMethod]
        public void Build_WithAsciiText_ReturnsSubsettedPrimaryFont()
        {
            var builder = new DocumentFontSubsetBuilder(TestFolderEngine);
            builder.AddText(TestFamily, FontSubFamily.Regular, "Hello");
            builder.Build();

            var provider = builder.GetShapingProvider(TestFamily, FontSubFamily.Regular);

            Assert.IsNotNull(provider);
            Assert.IsTrue(provider.PrimaryFont.IsSubset,
                "Ascii text through the primary font must yield a subsetted primary.");
        }

        [TestMethod]
        public void Build_WithMultipleAddTextCalls_CollectsAllCodePoints()
        {
            var builder = new DocumentFontSubsetBuilder(TestFolderEngine);
            builder.AddText(TestFamily, FontSubFamily.Regular, "abc");
            builder.AddText(TestFamily, FontSubFamily.Regular, "def");
            builder.Build();

            var provider = builder.GetShapingProvider(TestFamily, FontSubFamily.Regular);

            // Every code point from every AddText call must survive into the subset.
            foreach (var ch in "abcdef")
            {
                ushort glyphId;
                Assert.IsTrue(
                    provider.PrimaryFont.CmapTable.TryGetGlyphId(ch, out glyphId) && glyphId != 0,
                    "Subset must contain '" + ch + "' collected across multiple AddText calls.");
            }
        }

        [TestMethod]
        public void Build_WithEmoji_SubsetsFallbackFont()
        {
            var builder = new DocumentFontSubsetBuilder(TestFolderEngine);
            builder.AddText(TestFamily, FontSubFamily.Regular, char.ConvertFromUtf32(0x1F600)); // 😀
            builder.Build();

            // The emoji routes to the Noto Emoji fallback, which must appear among the embedded
            // fonts and carry the glyph.
            var embedded = builder.GetFontsToEmbed().ToList();

            bool emojiCarried = embedded.Any(sf =>
            {
                ushort glyphId;
                return sf.Font.CmapTable.TryGetGlyphId(0x1F600, out glyphId) && glyphId != 0;
            });

            Assert.IsTrue(emojiCarried, "The emoji fallback font must be subsetted and embedded.");
        }

        [TestMethod]
        public void Build_UnusedFallbackFontsAreExcluded()
        {
            // Pure ascii: only the primary is needed. No emoji/math fallback should be embedded.
            var builder = new DocumentFontSubsetBuilder(TestFolderEngine);
            builder.AddText(TestFamily, FontSubFamily.Regular, "Hello");
            builder.Build();

            var embedded = builder.GetFontsToEmbed().ToList();

            Assert.AreEqual(1, embedded.Count,
                "Only the primary font should be embedded when no fallback was needed.");
            StringAssert.Contains(embedded[0].Family, "Roboto");
        }
    }
}
