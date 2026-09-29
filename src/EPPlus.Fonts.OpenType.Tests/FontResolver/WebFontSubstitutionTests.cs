/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/28/2026         EPPlus Software AB           Web font substitution tests
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.FontResolver
{
    /// <summary>
    /// Tests for web font substitution via <see cref="OpenTypeFontEngine.GetFamilyForTarget"/>.
    /// None of these tests resolve or load fonts, so they are independent of installed fonts.
    /// </summary>
    [TestClass]
    public class WebFontSubstitutionTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        /// <summary>
        /// Creates an isolated engine for tests that change the configuration.
        /// The shared engines in <see cref="FontTestBase"/> must not be reconfigured.
        /// </summary>
        private static OpenTypeFontEngine CreateEngine(Action<IEpplusFontConfiguration>? configure = null)
        {
            return new OpenTypeFontEngine(cfg =>
            {
                cfg.SearchSystemDirectories = false;
                if (configure != null)
                {
                    configure(cfg);
                }
            });
        }

        // -----------------------------------------------------------------------------------------
        // Render target
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void DocumentTarget_NeverSubstitutes()
        {
            var family = TestFolderEngine.GetFamilyForTarget("Aptos Narrow", FontRenderTarget.Document);

            Assert.AreEqual("Aptos Narrow", family);
        }

        [TestMethod]
        public void WebTarget_UnknownFont_IsUnchanged()
        {
            var family = TestFolderEngine.GetFamilyForTarget("Arial", FontRenderTarget.Web);

            Assert.AreEqual("Arial", family);
        }

        [TestMethod]
        public void WebTarget_NullOrEmptyFontName_IsReturnedAsIs()
        {
            Assert.IsNull(TestFolderEngine.GetFamilyForTarget(null!, FontRenderTarget.Web));
            Assert.AreEqual("", TestFolderEngine.GetFamilyForTarget("", FontRenderTarget.Web));
        }

        // -----------------------------------------------------------------------------------------
        // Default substitution table
        // -----------------------------------------------------------------------------------------

        [DataTestMethod]
        [DataRow("Aptos", "Arial")]
        [DataRow("Aptos Narrow", "Calibri")]
        [DataRow("Aptos Display", "Arial")]
        [DataRow("Aptos Serif", "Cambria")]
        [DataRow("Aptos Mono", "Consolas")]
        [DataRow("Grandview", "Calibri")]
        [DataRow("Seaford", "Segoe UI")]
        [DataRow("Tenorite", "Segoe UI")]
        [DataRow("Bierstadt", "Arial")]
        [DataRow("Skeena", "Segoe UI")]
        public void WebTarget_OfficeCloudFont_IsSubstitutedByDefault(string requested, string expected)
        {
            var family = TestFolderEngine.GetFamilyForTarget(requested, FontRenderTarget.Web);

            Assert.AreEqual(expected, family);
        }

        [DataTestMethod]
        [DataRow("aptos narrow")]
        [DataRow("APTOS NARROW")]
        [DataRow("Aptos NARROW")]
        public void WebTarget_Lookup_IsCaseInsensitive(string requested)
        {
            var family = TestFolderEngine.GetFamilyForTarget(requested, FontRenderTarget.Web);

            Assert.AreEqual("Calibri", family);
        }

        // -----------------------------------------------------------------------------------------
        // User configuration
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void WebTarget_UserSubstitution_OverridesDefault()
        {
            using (var engine = CreateEngine(cfg => cfg.WebFontSubstitutions["Aptos Narrow"] = "Arial"))
            {
                Assert.AreEqual("Arial", engine.GetFamilyForTarget("Aptos Narrow", FontRenderTarget.Web));
            }
        }

        [TestMethod]
        public void WebTarget_UserSubstitution_ForNonDefaultFont_IsApplied()
        {
            using (var engine = CreateEngine(cfg => cfg.WebFontSubstitutions["Goudy Stout"] = "Georgia"))
            {
                Assert.AreEqual("Georgia", engine.GetFamilyForTarget("Goudy Stout", FontRenderTarget.Web));
            }
        }

        [TestMethod]
        public void WebTarget_UserSubstitution_DoesNotAffectDocumentTarget()
        {
            using (var engine = CreateEngine(cfg => cfg.WebFontSubstitutions["Goudy Stout"] = "Georgia"))
            {
                Assert.AreEqual("Goudy Stout", engine.GetFamilyForTarget("Goudy Stout", FontRenderTarget.Document));
            }
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow(null)]
        public void WebTarget_EmptySubstitute_KeepsOriginalFont(string? substitute)
        {
            using (var engine = CreateEngine(cfg => cfg.WebFontSubstitutions["Aptos Narrow"] = substitute!))
            {
                Assert.AreEqual("Aptos Narrow", engine.GetFamilyForTarget("Aptos Narrow", FontRenderTarget.Web));
            }
        }

        [TestMethod]
        public void WebTarget_RemovedDefault_KeepsOriginalFont()
        {
            using (var engine = CreateEngine(cfg => cfg.WebFontSubstitutions.Remove("Aptos Narrow")))
            {
                Assert.AreEqual("Aptos Narrow", engine.GetFamilyForTarget("Aptos Narrow", FontRenderTarget.Web));
            }
        }

        [TestMethod]
        public void WebTarget_Substitution_IsNotChained()
        {
            using (var engine = CreateEngine(cfg =>
            {
                cfg.WebFontSubstitutions["Font A"] = "Font B";
                cfg.WebFontSubstitutions["Font B"] = "Font C";
            }))
            {
                Assert.AreEqual("Font B", engine.GetFamilyForTarget("Font A", FontRenderTarget.Web));
            }
        }

        [TestMethod]
        public void Reset_RestoresDefaultSubstitutions()
        {
            using (var engine = CreateEngine(cfg =>
            {
                cfg.WebFontSubstitutions.Clear();
                cfg.WebFontSubstitutions["Goudy Stout"] = "Georgia";
                cfg.Reset();
            }))
            {
                Assert.AreEqual("Calibri", engine.GetFamilyForTarget("Aptos Narrow", FontRenderTarget.Web));
                Assert.AreEqual("Goudy Stout", engine.GetFamilyForTarget("Goudy Stout", FontRenderTarget.Web));
            }
        }

        // -----------------------------------------------------------------------------------------
        // Lifecycle
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        [ExpectedException(typeof(ObjectDisposedException))]
        public void GetFamilyForTarget_OnDisposedEngine_Throws()
        {
            var engine = CreateEngine();
            engine.Dispose();

            engine.GetFamilyForTarget("Aptos Narrow", FontRenderTarget.Web);
        }
    }
}