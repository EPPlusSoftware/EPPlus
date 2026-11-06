using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Fonts.OpenType;

namespace EPPlus.Export.Pdf.Tests
{
    [TestClass]
    public class PdfPageSettingsOrientationTests
    {
        private const double Tol = 1e-9;

        private static PdfPageSettings NewSettings() => new PdfPageSettings(new OpenTypeFontEngine());

        [TestMethod]
        public void Orientation_Landscape_MakesWidthExceedHeight_AndRecalculatesBounds()
        {
            var s = NewSettings();
            s.Orientation = Orientations.Landscape;

            Assert.IsTrue(s.PageSize.Width > s.PageSize.Height, "Landscape must have width > height.");
            Assert.AreEqual(
                s.PageSize.WidthPu - s.Margins.LeftPu - s.Margins.RightPu,
                s.ContentBounds.Width, Tol,
                "Setting the orientation must recompute the content bounds for the swapped page.");
        }

        [TestMethod]
        public void Orientation_BackToPortrait_SwapsBack()
        {
            var s = NewSettings();
            s.Orientation = Orientations.Landscape;
            s.Orientation = Orientations.Portrait;
            Assert.IsTrue(s.PageSize.Height >= s.PageSize.Width, "Portrait must have height >= width.");
        }

        [TestMethod]
        public void Margins_Setter_RecalculatesContentBounds()
        {
            var s = NewSettings();
            s.Margins = PdfMargins.Wide;
            Assert.AreEqual(PdfMargins.Wide.LeftPu, s.ContentBounds.Left, Tol);
        }
    }
}
