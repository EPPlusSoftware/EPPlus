/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           GsubFeature/GposFeature plumbing
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;
using System.Linq;

namespace EPPlus.Export.Pdf.Tests
{
    /// <summary>
    /// Covers PdfTextShaper.BuildShapingOptions, in particular GsubFeature.None/GposFeature.None.
    ///
    /// TextShaper.ApplyPositioning treats an empty or null GposFeatures list as "apply every GPOS
    /// feature" for kerning and mark positioning - so naively mapping GposFeature.None to an
    /// empty tag list and assigning it to ShapingOptions.GposFeatures would turn kerning and mark
    /// positioning ON, the opposite of what None means. BuildShapingOptions instead turns
    /// ApplyPositioning/ApplySubstitutions off outright when None is requested, which is
    /// unambiguous regardless of how an empty tag list is interpreted further down.
    /// </summary>
    [TestClass]
    public class PdfTextShaperShapingOptionsTests
    {
        public TestContext? TestContext { get; set; }

        private static PdfPageSettings CreateSettings()
        {
            // BuildShapingOptions never touches pageSettings.FontEngine, but the engine is kept
            // alive (not disposed) regardless, so this helper stays valid if that ever changes.
            var engine = new OpenTypeFontEngine();
            return new PdfPageSettings(engine);
        }

        [TestMethod]
        public void BuildShapingOptions_Default_AppliesBothAndUsesDefaultTags()
        {
            var settings = CreateSettings();
            // Defaults per PdfPageSettings: GsubFeatures = Liga|Clig, GposFeatures = Kern|Mark.

            var options = PdfTextShaper.BuildShapingOptions(settings);

            Assert.IsTrue(options.ApplySubstitutions);
            Assert.IsTrue(options.ApplyPositioning);
            CollectionAssert.AreEquivalent(new[] { "liga", "clig" }, options.GsubFeatures.ToList());
            CollectionAssert.AreEquivalent(new[] { "kern", "mark" }, options.GposFeatures.ToList());
        }

        [TestMethod]
        public void BuildShapingOptions_GsubNone_TurnsOffSubstitutionsEntirely()
        {
            var settings = CreateSettings();
            settings.GsubFeatures = GsubFeature.None;

            var options = PdfTextShaper.BuildShapingOptions(settings);

            Assert.IsFalse(
                options.ApplySubstitutions,
                "GsubFeature.None must disable substitutions outright, not rely on an empty tag "
                + "list being interpreted as \"apply nothing\" somewhere downstream");
        }

        [TestMethod]
        public void BuildShapingOptions_GposNone_TurnsOffPositioningEntirely()
        {
            var settings = CreateSettings();
            settings.GposFeatures = GposFeature.None;

            var options = PdfTextShaper.BuildShapingOptions(settings);

            Assert.IsFalse(
                options.ApplyPositioning,
                "GposFeature.None must disable positioning outright. TextShaper.ApplyPositioning "
                + "treats an empty/null GposFeatures list as \"apply every feature\" for kerning "
                + "and mark, so leaving ApplyPositioning on with an empty list here would turn "
                + "kerning and mark ON instead of off.");
        }

        [TestMethod]
        public void BuildShapingOptions_GsubDligOnly_MapsToSingleTag()
        {
            var settings = CreateSettings();
            settings.GsubFeatures = GsubFeature.Dlig;

            var options = PdfTextShaper.BuildShapingOptions(settings);

            Assert.IsTrue(options.ApplySubstitutions);
            CollectionAssert.AreEquivalent(new[] { "dlig" }, options.GsubFeatures.ToList());
        }
    }
}