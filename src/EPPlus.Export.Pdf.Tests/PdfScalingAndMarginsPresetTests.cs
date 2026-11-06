using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Graphics.Units;

namespace EPPlus.Export.Pdf.Tests
{
    [TestClass]
    public class PdfScalingAndMarginsPresetTests
    {
        private const double Tol = 1e-9;

        [TestMethod]
        public void Scaling_ScaleCtor_UsesAdjustToNormalSize()
        {
            var sc = new PdfScaling(0.75);
            Assert.AreEqual(ScalingMode.AdjustToNormalSize, sc.ScalingMode);
            Assert.AreEqual(0.75, sc.Scale, Tol);
        }

        [TestMethod]
        public void Scaling_PagesCtor_UsesFitToPages()
        {
            var sc = new PdfScaling(2, 3);
            Assert.AreEqual(ScalingMode.FitToPages, sc.ScalingMode);
            Assert.AreEqual(2, sc.PagesWide);
            Assert.AreEqual(3, sc.PagesTall);
        }

        [TestMethod]
        public void Scaling_Presets_HaveExpectedSemantics()
        {
            Assert.AreEqual(ScalingMode.AdjustToNormalSize, PdfScaling.NoScaling.ScalingMode);
            Assert.AreEqual(1d, PdfScaling.NoScaling.Scale, Tol);

            Assert.AreEqual(ScalingMode.FitToPages, PdfScaling.FitSheetToOnePage.ScalingMode);
            Assert.AreEqual(1, PdfScaling.FitSheetToOnePage.PagesWide);
            Assert.AreEqual(1, PdfScaling.FitSheetToOnePage.PagesTall);

            // "All columns on one page" => 1 wide, unconstrained tall (0).
            Assert.AreEqual(1, PdfScaling.FitAllColumnsOnOnePage.PagesWide);
            Assert.AreEqual(0, PdfScaling.FitAllColumnsOnOnePage.PagesTall);

            // "All rows on one page" => unconstrained wide (0), 1 tall.
            Assert.AreEqual(0, PdfScaling.FitAllRowsOnOnePage.PagesWide);
            Assert.AreEqual(1, PdfScaling.FitAllRowsOnOnePage.PagesTall);
        }

        [TestMethod]
        public void Margins_Normal_ConvertsMmToPoints()
        {
            var m = PdfMargins.Normal;
            Assert.AreEqual(UnitConversion.MmToPoints(m.Top), m.TopPu, Tol);
            Assert.AreEqual(UnitConversion.MmToPoints(m.Left), m.LeftPu, Tol);
            Assert.AreEqual(UnitConversion.MmToPoints(m.Right), m.RightPu, Tol);
            Assert.AreEqual(UnitConversion.MmToPoints(m.Bottom), m.BottomPu, Tol);
            Assert.AreEqual(UnitConversion.MmToPoints(m.Header), m.HeaderPu, Tol);
            Assert.AreEqual(UnitConversion.MmToPoints(m.Footer), m.FooterPu, Tol);
        }

        [TestMethod]
        public void Margins_Wide_IsLargerThanNarrow()
        {
            Assert.IsTrue(PdfMargins.Wide.LeftPu > PdfMargins.Narrow.LeftPu);
            Assert.IsTrue(PdfMargins.Wide.TopPu > PdfMargins.Narrow.TopPu);
        }
    }
}
