using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tests.Tables;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class WebEmbeddingPolicyTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }
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

        // ---- Level 1b: GetEmbeddingPermission / IsBitmapEmbeddingOnly (pure fsType interpretation) ----

        [TestMethod]
        [DataRow(FsTypeFlags.Installable, FontEmbeddingPermission.Installable)]
        [DataRow(FsTypeFlags.RestrictedLicense, FontEmbeddingPermission.RestrictedLicense)]
        [DataRow(FsTypeFlags.PreviewPrint, FontEmbeddingPermission.PreviewAndPrint)]
        [DataRow(FsTypeFlags.Editable, FontEmbeddingPermission.Editable)]
        public void GetEmbeddingPermission_SingleUsageBit_ReturnsMatchingPermission(
            FsTypeFlags fsType, FontEmbeddingPermission expected)
        {
            var os2 = new Os2Table { fsType = fsType };
            Assert.AreEqual(expected, os2.GetEmbeddingPermission());
        }

        [TestMethod]
        [DataRow(FsTypeFlags.RestrictedLicense | FsTypeFlags.PreviewPrint, FontEmbeddingPermission.PreviewAndPrint)]
        [DataRow(FsTypeFlags.RestrictedLicense | FsTypeFlags.Editable, FontEmbeddingPermission.Editable)]
        [DataRow(FsTypeFlags.PreviewPrint | FsTypeFlags.Editable, FontEmbeddingPermission.Editable)]
        [DataRow(FsTypeFlags.RestrictedLicense | FsTypeFlags.PreviewPrint | FsTypeFlags.Editable, FontEmbeddingPermission.Editable)]
        public void GetEmbeddingPermission_SeveralUsageBits_LeastRestrictiveWins(
            FsTypeFlags fsType, FontEmbeddingPermission expected)
        {
            var os2 = new Os2Table { fsType = fsType };
            Assert.AreEqual(expected, os2.GetEmbeddingPermission());
        }

        [TestMethod]
        [DataRow(FsTypeFlags.NoSubsetting | FsTypeFlags.BitmapEmbeddingOnly, FontEmbeddingPermission.Installable)]
        [DataRow(FsTypeFlags.PreviewPrint | FsTypeFlags.NoSubsetting | FsTypeFlags.BitmapEmbeddingOnly, FontEmbeddingPermission.PreviewAndPrint)]
        public void GetEmbeddingPermission_HighBits_DoNotAffectUsageLevel(
            FsTypeFlags fsType, FontEmbeddingPermission expected)
        {
            var os2 = new Os2Table { fsType = fsType };
            Assert.AreEqual(expected, os2.GetEmbeddingPermission());
        }

        [TestMethod]
        [DataRow(FsTypeFlags.BitmapEmbeddingOnly, true)]
        [DataRow(FsTypeFlags.BitmapEmbeddingOnly | FsTypeFlags.Editable, true)]
        [DataRow(FsTypeFlags.NoSubsetting, false)]
        [DataRow(FsTypeFlags.Installable, false)]
        public void IsBitmapEmbeddingOnly_ReflectsBit9(FsTypeFlags fsType, bool expected)
        {
            var os2 = new Os2Table { fsType = fsType };
            Assert.AreEqual(expected, os2.IsBitmapEmbeddingOnly());
        }

        // ---- Level 2b: web default policy (no callback) ----

        [TestMethod]
        [DataRow(FsTypeFlags.Installable)]
        [DataRow(FsTypeFlags.Editable)]
        public void ResolveEmbeddingDecision_Web_PermissiveFsType_ReturnsSubset(FsTypeFlags fsType)
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = fsType;
            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        [DataRow(FsTypeFlags.RestrictedLicense)]
        [DataRow(FsTypeFlags.PreviewPrint)]
        [DataRow(FsTypeFlags.NoSubsetting)]
        [DataRow(FsTypeFlags.BitmapEmbeddingOnly)]
        [DataRow(FsTypeFlags.Editable | FsTypeFlags.NoSubsetting)]
        [DataRow(FsTypeFlags.Editable | FsTypeFlags.BitmapEmbeddingOnly)]
        public void ResolveEmbeddingDecision_Web_RestrictiveFsType_ReturnsSkipWithoutThrowing(FsTypeFlags fsType)
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = fsType;
            Assert.AreEqual(FontEmbeddingDecision.Skip,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_BundledFont_ReturnsSubsetRegardlessOfFsType()
        {
            // Bundled fonts are OFL-licensed; the bundled check runs before fsType is consulted.
            var font = TestFolderEngine.LoadFont("Archivo Narrow", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.RestrictedLicense;
            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        // ---- Level 2c: document (PDF) path is unchanged ----

        [TestMethod]
        [DataRow(FsTypeFlags.Installable)]
        [DataRow(FsTypeFlags.PreviewPrint)]
        [DataRow(FsTypeFlags.Editable)]
        [DataRow(FsTypeFlags.NoSubsetting)]
        [DataRow(FsTypeFlags.BitmapEmbeddingOnly)]
        public void ResolveEmbeddingDecision_OverloadWithoutTarget_EqualsDocument(FsTypeFlags fsType)
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = fsType;
            Assert.AreEqual(
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Document),
                TestFolderEngine.ResolveEmbeddingDecision(font));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Document_PreviewPrint_ReturnsSubset()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.PreviewPrint;
            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Document));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Document_BitmapEmbeddingOnly_IsIgnored()
        {
            // Documents current behaviour: the PDF path does not act on bitmap-only. Changing that is
            // a separate decision; if it is made, this test is expected to change with it.
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.BitmapEmbeddingOnly;
            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Document));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Document_RestrictedLicense_StillThrows()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.RestrictedLicense;
            Assert.ThrowsExactly<InvalidOperationException>(
                () => TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Document));
        }

        // ---- Level 2d: web default policy requires an open licence ----

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_NoLicenseRecords_ReturnsSkip()
        {
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.Installable;
            LicenseDetectionTests.SetLicenseRecords(font, null, null);

            Assert.AreEqual(FontEmbeddingDecision.Skip,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        [DataRow(FsTypeFlags.Installable)]
        [DataRow(FsTypeFlags.Editable)]
        public void ResolveEmbeddingDecision_Web_CommercialLicense_ReturnsSkipWhateverFsType(FsTypeFlags fsType)
        {
            // fsType Editable permits embedding in documents, not use as a web font.
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = fsType;
            LicenseDetectionTests.SetLicenseRecords(font, "Commercial font. All rights reserved.", "https://example.com/eula");

            Assert.AreEqual(FontEmbeddingDecision.Skip,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_OpenLicenseButRestrictiveFsType_ReturnsSkip()
        {
            // An open licence is not enough on its own; fsType must also permit embedding.
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.PreviewPrint;
            Assert.IsTrue(font.NameTable.HasOpenLicense(), "Test setup: Roboto is Apache licensed.");

            Assert.AreEqual(FontEmbeddingDecision.Skip,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_NoLicense_CallbackSubset_Wins()
        {
            var engine = CreateEngineWithCallback(info => FontEmbeddingDecision.Subset);
            var font = engine.LoadFont("Roboto", ignoreCache: true);
            LicenseDetectionTests.SetLicenseRecords(font, null, null);

            Assert.AreEqual(FontEmbeddingDecision.Subset,
                engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_BundledFont_WithoutLicenseRecords_ReturnsSubset()
        {
            var font = TestFolderEngine.LoadFont("Archivo Narrow", ignoreCache: true);
            LicenseDetectionTests.SetLicenseRecords(font, null, null);

            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Document_NoLicense_IsUnchanged()
        {
            // The licence rule applies to web output only.
            var font = TestFolderEngine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.Installable;
            LicenseDetectionTests.SetLicenseRecords(font, null, null);

            Assert.AreEqual(FontEmbeddingDecision.Subset,
                TestFolderEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Document));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_CallbackReceivesLicense_OpenFont()
        {
            FontEmbeddingInfo captured = null;
            var engine = CreateEngineWithCallback(info =>
            {
                captured = info;
                return FontEmbeddingDecision.Default;
            });

            var font = engine.LoadFont("Roboto", ignoreCache: true);
            engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web);

            Assert.IsNotNull(captured);
            Assert.IsTrue(captured.HasOpenLicense);
            StringAssert.Contains(captured.LicenseDescription, "Apache License");
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_CallbackReceivesLicense_NoLicense()
        {
            FontEmbeddingInfo captured = null;
            var engine = CreateEngineWithCallback(info =>
            {
                captured = info;
                return FontEmbeddingDecision.Default;
            });

            var font = engine.LoadFont("Roboto", ignoreCache: true);
            LicenseDetectionTests.SetLicenseRecords(font, null, null);
            engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web);

            Assert.IsNotNull(captured);
            Assert.IsFalse(captured.HasOpenLicense);
            Assert.IsNull(captured.LicenseDescription);
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_SystemCommercialFont_ReturnsSkip()
        {
            // Calibri is a Microsoft font with fsType Editable; it must not be served as a web font.
            RequireFont(SystemFontsEngine, "Calibri");
            var font = SystemFontsEngine.LoadFont("Calibri");

            Assert.IsFalse(font.NameTable.HasOpenLicense());
            Assert.AreEqual(FontEmbeddingDecision.Skip,
                SystemFontsEngine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        // ---- Level 3b: callback with target ----

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_CallbackReceivesTargetPermissionAndBitmapFlag()
        {
            FontEmbeddingInfo captured = null;
            var engine = CreateEngineWithCallback(info =>
            {
                captured = info;
                return FontEmbeddingDecision.Default;
            });

            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.Editable | FsTypeFlags.BitmapEmbeddingOnly;

            engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web);

            Assert.IsNotNull(captured);
            Assert.AreEqual(FontRenderTarget.Web, captured.Target);
            Assert.AreEqual(FontEmbeddingPermission.Editable, captured.Permission);
            Assert.IsTrue(captured.BitmapEmbeddingOnly);
            Assert.AreEqual(FontEmbeddingRestriction.None, captured.Restriction);
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Document_CallbackReceivesDocumentTarget()
        {
            FontEmbeddingInfo captured = null;
            var engine = CreateEngineWithCallback(info =>
            {
                captured = info;
                return FontEmbeddingDecision.Default;
            });

            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.PreviewPrint;

            engine.ResolveEmbeddingDecision(font);

            Assert.IsNotNull(captured);
            Assert.AreEqual(FontRenderTarget.Document, captured.Target);
            Assert.AreEqual(FontEmbeddingPermission.PreviewAndPrint, captured.Permission);
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_CallbackSubset_OverridesWebPolicy()
        {
            // A caller holding a web-font licence can opt in to embedding a Preview & Print font.
            var engine = CreateEngineWithCallback(info => FontEmbeddingDecision.Subset);
            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.PreviewPrint;

            Assert.AreEqual(FontEmbeddingDecision.Subset,
                engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        [TestMethod]
        public void ResolveEmbeddingDecision_Web_CallbackDefault_FallsThroughToWebPolicy()
        {
            var engine = CreateEngineWithCallback(info => FontEmbeddingDecision.Default);
            var font = engine.LoadFont("Roboto", ignoreCache: true);
            font.Os2Table.fsType = FsTypeFlags.PreviewPrint;

            Assert.AreEqual(FontEmbeddingDecision.Skip,
                engine.ResolveEmbeddingDecision(font, FontRenderTarget.Web));
        }

        // ---- FontEmbeddingInfo backward compatibility ----

        [TestMethod]
        public void FontEmbeddingInfo_TwoArgumentConstructor_DefaultsToInstallableDocument()
        {
            var info = new FontEmbeddingInfo("Some Font", FontEmbeddingRestriction.None);

            Assert.AreEqual(FontEmbeddingPermission.Installable, info.Permission);
            Assert.IsFalse(info.BitmapEmbeddingOnly);
            Assert.AreEqual(FontRenderTarget.Document, info.Target);
        }

        [TestMethod]
        public void FontEmbeddingInfo_FiveArgumentConstructor_HasNoOpenLicense()
        {
            var info = new FontEmbeddingInfo("Some Font", FontEmbeddingRestriction.None,
                FontEmbeddingPermission.Editable, false, FontRenderTarget.Web);

            Assert.IsFalse(info.HasOpenLicense);
            Assert.IsNull(info.LicenseDescription);
        }
    }
}
