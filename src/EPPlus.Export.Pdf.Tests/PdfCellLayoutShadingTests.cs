using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Fonts.OpenType;

namespace EPPlus.Export.Pdf.Tests
{
    [TestClass]
    public class PdfCellLayoutShadingTests
    {
        private const double Tol = 1e-9;

        private static PdfPageSettings NewSettings() => new PdfPageSettings(new OpenTypeFontEngine());

        private static PdfCellLayout LinearCell(double degree)
        {
            var cell = new PdfCellLayout(10, 100, 50, 20); // x, y, width, height
            cell.CellFillData.GradientFillData = new PdfCellGradientFillData
            {
                GradientType = ExcelFillGradientType.Linear,
                Degree = degree
            };
            return cell;
        }

        [TestMethod]
        public void UpdateShadingPositionMatrix_Linear_MapsDegreeToCoords()
        {
            var expected = new Dictionary<double, double[]>
            {
                { 45d,  new double[] { 0, 1, 1, 0 } },
                { 90d,  new double[] { 0, 1, 0, 0 } },
                { 135d, new double[] { 1, 1, 0, 0 } },
                { 180d, new double[] { 1, 0, 0, 0 } },
                { 225d, new double[] { 1, 0, 0, 1 } },
                { 270d, new double[] { 0, 0, 0, 1 } },
                { 315d, new double[] { 0, 0, 1, 1 } },
                { 0d,   new double[] { 0, 0, 1, 0 } },
                { 12.5d /* unmapped -> default */, new double[] { 0, 0, 1, 0 } },
            };

            foreach (var kv in expected)
            {
                var cell = LinearCell(kv.Key);
                cell.UpdateShadingPositionMatrix(NewSettings());
                CollectionAssert.AreEqual(kv.Value, cell.CellFillData.GradientFillData.coords,
                    $"coords for degree {kv.Key}");
            }
        }

        [TestMethod]
        public void UpdateShadingPositionMatrix_UnitScale_MatrixIsSizeAndPosition()
        {
            var cell = LinearCell(0d);                 // Size=(50,20), LocalPosition=(10, 100-20=80)
            cell.UpdateShadingPositionMatrix(NewSettings()); // NoScaling -> ContentScale = 1

            CollectionAssert.AreEqual(
                new double[] { 50, 0, 0, 20, 10, 80 },
                cell.CellFillData.GradientFillData.matrix);
        }

        [TestMethod]
        public void UpdateShadingPositionMatrix_HalfScale_AppliesAnchorOffset()
        {
            const double s = 0.5d;
            var settings = NewSettings();
            settings.Scaling = new PdfScaling(s);      // AdjustToNormalSize, Scale = 0.5
            double ax = settings.ContentBounds.Left;
            double ay = settings.ContentBounds.Top;

            var cell = LinearCell(0d);                 // Size=(50,20), LocalPosition=(10,80)
            cell.UpdateShadingPositionMatrix(settings);

            var m = cell.CellFillData.GradientFillData.matrix;
            Assert.AreEqual(50 * s, m[0], Tol);
            Assert.AreEqual(0d, m[1], Tol);
            Assert.AreEqual(0d, m[2], Tol);
            Assert.AreEqual(20 * s, m[3], Tol);
            Assert.AreEqual(10 * s + (1d - s) * ax, m[4], Tol);
            Assert.AreEqual(80 * s + (1d - s) * ay, m[5], Tol);
        }

        [TestMethod]
        public void UpdateShadingPositionMatrix_PathNoEdges_MapsToCornerCoords()
        {
            var cell = new PdfCellLayout(10, 100, 50, 20);
            cell.CellFillData.GradientFillData = new PdfCellGradientFillData
            {
                GradientType = ExcelFillGradientType.Path,
                Top = 0,
                Bottom = 0,
                Left = 0,
                Right = 0
            };

            cell.UpdateShadingPositionMatrix(NewSettings());

            CollectionAssert.AreEqual(
                new double[] { 0, 1, 0, 0, 1, 1 },
                cell.CellFillData.GradientFillData.coords);
        }

        [TestMethod]
        public void UpdateShadingPositionMatrix_NoGradient_DoesNothing()
        {
            var cell = new PdfCellLayout(10, 100, 50, 20); // GradientFillData stays null
            cell.UpdateShadingPositionMatrix(NewSettings());
            Assert.IsNull(cell.CellFillData.GradientFillData);
        }
    }
}
