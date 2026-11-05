using EPPlus.Export.Pdf.Settings.PdfPageSizes;

namespace EPPlus.Export.Pdf.Tests
{
    [TestClass]
    public class PdfContentBoundsTests
    {
        private const double Tol = 1e-9;

        [TestMethod]
        public void CalculateBounds_A4Normal_ComputesRectAndHeaderFooterThirds()
        {
            var m = PdfMargins.Normal;
            var size = PdfPageSize.A4;

            var b = new PdfContentBounds(m, size); // ctor calls CalculateBounds

            Assert.AreEqual(size.WidthPu - m.LeftPu - m.RightPu, b.Width, Tol);
            Assert.AreEqual(m.LeftPu, b.Left, Tol);
            Assert.AreEqual(b.Left + b.Width, b.Right, Tol);
            Assert.AreEqual(size.HeightPu - m.TopPu, b.Top, Tol);
            Assert.AreEqual(m.BottomPu, b.Bottom, Tol);
            Assert.AreEqual(b.Top - b.Bottom, b.Height, Tol);

            var third = b.Width / 3d;
            Assert.AreEqual(b.Left + third, b.CenterHeaderX, Tol);
            Assert.AreEqual(b.Left + 2 * third, b.RightHeaderX, Tol);
            Assert.AreEqual(b.Left + third, b.CenterFooterX, Tol);
            Assert.AreEqual(b.Left + 2 * third, b.RightFooterX, Tol);
            Assert.AreEqual(size.HeightPu - m.HeaderPu, b.HeaderY, Tol);
            Assert.AreEqual(m.FooterPu, b.FooterY, Tol);
        }

        [TestMethod]
        public void CalculateBounds_WideMargins_LeavesSmallerContentArea()
        {
            var normal = new PdfContentBounds(PdfMargins.Normal, PdfPageSize.A4);
            var wide = new PdfContentBounds(PdfMargins.Wide, PdfPageSize.A4);

            Assert.IsTrue(wide.Width < normal.Width, "Wider margins must shrink the content width.");
            Assert.IsTrue(wide.Height < normal.Height, "Wider margins must shrink the content height.");
        }
    }
}
