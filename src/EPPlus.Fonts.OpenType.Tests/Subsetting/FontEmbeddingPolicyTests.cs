using EPPlus.Fonts.OpenType.Tables.Os2;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class FontEmbeddingPolicyTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void GetEmbeddingRestriction_Installable_ReturnsNone()
        {
            var os2 = new Os2Table { fsType = FsTypeFlags.Installable };
            Assert.AreEqual(FontEmbeddingRestriction.None, os2.GetEmbeddingRestriction());
        }

        [TestMethod]
        public void GetEmbeddingRestriction_RestrictedLicense_ReturnsNoEmbedding()
        {
            var os2 = new Os2Table { fsType = FsTypeFlags.RestrictedLicense };
            Assert.AreEqual(FontEmbeddingRestriction.NoEmbedding, os2.GetEmbeddingRestriction());
        }

        [TestMethod]
        public void GetEmbeddingRestriction_NoSubsetting_ReturnsNoSubsetting()
        {
            var os2 = new Os2Table { fsType = FsTypeFlags.NoSubsetting };
            Assert.AreEqual(FontEmbeddingRestriction.NoSubsetting, os2.GetEmbeddingRestriction());
        }

        [TestMethod]
        public void GetEmbeddingRestriction_RestrictedPlusNoSubsetting_NoEmbeddingWins()
        {
            var os2 = new Os2Table { fsType = FsTypeFlags.RestrictedLicense | FsTypeFlags.NoSubsetting };
            Assert.AreEqual(FontEmbeddingRestriction.NoEmbedding, os2.GetEmbeddingRestriction());
        }

        [TestMethod]
        public void GetEmbeddingRestriction_PreviewPrint_ReturnsNone()
        {
            var os2 = new Os2Table { fsType = FsTypeFlags.PreviewPrint };
            Assert.AreEqual(FontEmbeddingRestriction.None, os2.GetEmbeddingRestriction());
        }

        // ---- Level 2: ResolveEmbeddingDecision (policy + callback) ----
        // Uses Roboto and mutates fsType. Roboto itself is Installable, so the
        // baseline decision without mutation is Subset.

        [TestMethod]
        public void ResolveEmbeddingDecision_MutationPersists()
        {
            // Guards the whole level-2 suite: if mutating fsType on a loaded font
            // did not stick, every test below would be a false pass.
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.RestrictedLicense;
            Assert.AreEqual(FsTypeFlags.RestrictedLicense, font.Os2Table.fsType);
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Installable_NoCallback_ReturnsSubset()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.Installable;
            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_NoSubsetting_NoCallback_ReturnsEmbedWhole()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.NoSubsetting;
            Assert.AreEqual(FontEmbeddingDecision.EmbedWhole,
                TestFolderEngine.ResolveEmbeddingDecision(font));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_RestrictedLicense_NoCallback_Throws()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.RestrictedLicense;
            Assert.ThrowsExactly<InvalidOperationException>(
                () => TestFolderEngine.ResolveEmbeddingDecision(font));
        }

        // ---- Level 3: callback override ----
        // TestFolderEngine's config is locked at construction, so callback tests
        // build their own engine with the same font folders plus OnFontEmbedding.

        private static OpenTypeFontEngine CreateEngineWithCallback(
            Func<FontEmbeddingInfo, FontEmbeddingDecision> callback)
        {
            return new OpenTypeFontEngine(cfg =>
            {
                foreach (var folder in FontFolders)
                    cfg.FontDirectories.Add(folder);
                cfg.SearchSystemDirectories = false;
                cfg.OnFontEmbedding(callback);
            });
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_RestrictedLicense_CallbackSubset_OverridesAndDoesNotThrow()
        {
            var engine = CreateEngineWithCallback(info => FontEmbeddingDecision.Subset);
            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.RestrictedLicense;

            Assert.AreEqual(FontEmbeddingDecision.Subset,
                engine.ResolveEmbeddingDecision(font));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_RestrictedLicense_CallbackDefault_FallsThroughToPolicyAndThrows()
        {
            var engine = CreateEngineWithCallback(info => FontEmbeddingDecision.Default);
            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.RestrictedLicense;

            Assert.ThrowsExactly<InvalidOperationException>(
                () => engine.ResolveEmbeddingDecision(font));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_CallbackReceivesCorrectInfo()
        {
            FontEmbeddingInfo captured = null;
            var engine = CreateEngineWithCallback(info =>
            {
                captured = info;
                return FontEmbeddingDecision.Default;
            });

            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.NoSubsetting;

            // NoSubsetting + Default falls through to EmbedWhole (no throw), so this is safe to call.
            engine.ResolveEmbeddingDecision(font);

            Assert.IsNotNull(captured);
            Assert.AreEqual(FontEmbeddingRestriction.NoSubsetting, captured.Restriction);
            StringAssert.Contains(captured.FontName, "Roboto");
        }

        [TestMethod]
        public void CreateSubsettedProvider_NoSubsettingFont_EmbedsWholeFontNotSubset()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.NoSubsetting;

            var manager = new FontSubsetManager(TestFolderEngine, font);
            // Collect some code points so the font would otherwise be subsetted.
            manager.AddText("Hello");   // <-- vet ej exakt API-namn, se nedan

            var provider = manager.CreateSubsettedProvider();

            Assert.IsFalse(provider.PrimaryFont.IsSubset,
                "NoSubsetting font must be embedded whole, not subsetted.");
        }

        // -----------------------------------------------------------------------------------------
        // Level 4: Skip as a real fallback path (colleague feedback).
        //
        // A Skip decision can only originate from the OnFontEmbedding callback — the fsType policy
        // never produces it. When a font is skipped it must be removed from the effective chain and
        // its code points redistributed over the remaining fonts, rather than throwing. These tests
        // build an engine whose callback skips a specific font by name.
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void CreateSubsettedProvider_SkippedPrimary_NextFontBecomesPrimary()
        {
            // Roboto is the primary; the callback skips it. The provider's default fallback chain
            // (Noto Emoji, Noto Math) plus the resolver's last resort should take over, so the
            // resulting primary must be something other than Roboto and must not be null.
            var engine = CreateEngineWithCallback(info =>
                info.FontName != null && info.FontName.Contains("Roboto")
                    ? FontEmbeddingDecision.Skip
                    : FontEmbeddingDecision.Default);

            var roboto = engine.LoadFont("Roboto", ignoreCache: true);

            var manager = new FontSubsetManager(engine, roboto);
            manager.AddText("Hello");

            var provider = manager.CreateSubsettedProvider();

            Assert.IsNotNull(provider.PrimaryFont);
            StringAssert.DoesNotMatch(
                provider.PrimaryFont.GetEnglishFontFamilyName(),
                new System.Text.RegularExpressions.Regex("Roboto"),
                "A skipped primary must not remain the provider's primary font.");
        }

        [TestMethod]
        public void CreateSubsettedProvider_SkippedPrimary_AllTextSkipped_UsesLastResort()
        {
            // With ONLY Latin text and Roboto skipped, none of the default fallbacks (Emoji, Math)
            // cover the letters. The chain would collapse to empty, so the last-resort font
            // (Archivo Narrow) must step in and carry the glyphs.
            var engine = CreateEngineWithCallback(info =>
                info.FontName != null && info.FontName.Contains("Roboto")
                    ? FontEmbeddingDecision.Skip
                    : FontEmbeddingDecision.Default);

            var roboto = engine.LoadFont("Roboto", ignoreCache: true);

            var manager = new FontSubsetManager(engine, roboto);
            manager.AddText("Hello");

            var provider = manager.CreateSubsettedProvider();

            // Archivo Narrow is the guaranteed last resort. Its family name identifies it.
            StringAssert.Contains(
                provider.PrimaryFont.GetEnglishFontFamilyName(),
                "Archivo",
                "When the whole chain is skipped, the last-resort font must become primary.");

            // The redistributed Latin code points must actually be present in that font.
            ushort glyphId;
            Assert.IsTrue(
                provider.PrimaryFont.CmapTable.TryGetGlyphId('H', out glyphId) && glyphId != 0,
                "Latin glyphs must be carried by the last-resort font after redistribution.");

        }

        [TestMethod]
        public void CreateSubsettedProvider_SkippedPrimary_CjkText_GlyphsLandInReplacement()
        {
            // The heart of the redistribution logic: Roboto (Latin) is primary and covers none of
            // the CJK text. A CJK-capable fallback (BIZ UDGothic) sits in a CustomFontProvider chain.
            // When Roboto is skipped, the CJK code points that were distributed to it must be
            // redistributed to BIZ UDGothic and appear in the subsetted result.
            var engine = CreateEngineWithCallback(info =>
                info.FontName != null && info.FontName.Contains("Roboto")
                    ? FontEmbeddingDecision.Skip
                    : FontEmbeddingDecision.Default);

            var roboto = engine.LoadFont("Roboto", ignoreCache: true);
            var biz = engine.LoadFont("BIZ UDGothic", ignoreCache: true);

            var source = new CustomFontProvider(roboto);
            source.AddFallback(biz);

            var manager = new FontSubsetManager(engine, source);

            // U+6F22 漢 — a Han ideograph covered by BIZ UDGothic, not by Roboto.
            const int han = 0x6F22;
            manager.AddText(char.ConvertFromUtf32(han));

            var provider = manager.CreateSubsettedProvider();

            // Roboto skipped → the CJK-capable font becomes primary.
            ushort glyphId;
            Assert.IsTrue(
                provider.PrimaryFont.CmapTable.TryGetGlyphId((uint)han, out glyphId) && glyphId != 0,
                "The Han code point must be carried (and subsetted) by the replacement font.");

            StringAssert.DoesNotMatch(
                provider.PrimaryFont.GetEnglishFontFamilyName(),
                new System.Text.RegularExpressions.Regex("Roboto"),
                "The skipped primary must not remain the provider's primary font.");

            StringAssert.Contains(
                provider.PrimaryFont.GetEnglishFontFamilyName(),
                "BIZ",
                "The CJK-capable fallback must have become the primary font.");
        }

        [TestMethod]
        public void CreateSubsettedProvider_SkippedPrimary_PrefersChainFontOverLastResort()
        {
            // A skipped primary must hand off to a real font from the chain, NOT jump straight
            // to the Archivo Narrow last resort. Roboto (Latin) is primary; BIZ UDGothic is a
            // fallback that covers the CJK text. When Roboto is skipped, BIZ — not Archivo —
            // must become primary.
            var engine = CreateEngineWithCallback(info =>
                info.FontName != null && info.FontName.Contains("Roboto")
                    ? FontEmbeddingDecision.Skip
                    : FontEmbeddingDecision.Default);

            var roboto = engine.LoadFont("Roboto", ignoreCache: true);
            var biz = engine.LoadFont("BIZ UDGothic", ignoreCache: true);

            var source = new CustomFontProvider(roboto);
            source.AddFallback(biz);

            var manager = new FontSubsetManager(engine, source);
            manager.AddText(char.ConvertFromUtf32(0x6F22)); // 漢

            var provider = manager.CreateSubsettedProvider();

            var family = provider.PrimaryFont.GetEnglishFontFamilyName();

            // The positive assertion: the chain font took over.
            StringAssert.Contains(family, "BIZ",
                "A chain fallback must take over a skipped primary.");

            // The negative assertion — the crux: the last resort was NOT used.
            StringAssert.DoesNotMatch(
                family,
                new System.Text.RegularExpressions.Regex("Archivo"),
                "The last-resort font must not pre-empt an available chain fallback.");
        }
    }
}
