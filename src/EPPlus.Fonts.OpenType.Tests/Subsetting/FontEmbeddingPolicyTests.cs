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
    }
}
