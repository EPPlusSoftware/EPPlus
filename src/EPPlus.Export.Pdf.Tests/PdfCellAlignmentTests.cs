/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************/
using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Layout;
using Vector2 = EPPlus.Graphics.Geometry.Vector2;

namespace EPPlus.Export.Pdf.Tests
{
    /// <summary>
    /// Unit tests for <see cref="PdfCellAlignment"/> - the text-placement math extracted from
    /// PdfCellContentLayout. These assert the horizontal/vertical alignment rules, the
    /// "General" numeric-vs-text rule, the rotation offset, and the rotated bounding-box
    /// placement, all without constructing a full PdfCellContentLayout.
    ///
    /// Expected values are written as literals; the comment next to each one shows how it
    /// follows from the inputs, so the number is the assertion and the arithmetic is just
    /// documentation.
    /// </summary>
    [TestClass]
    public class PdfCellAlignmentTests
    {
        private const double RightMargin = 1.4d;
        private const double BottomMargin = 3.5d;
        private const double Tol = 1e-9;

        private static PdfCellAlignmentData Align(
            ExcelHorizontalAlignment h = ExcelHorizontalAlignment.General,
            ExcelVerticalAlignment v = ExcelVerticalAlignment.Bottom,
            int rotation = 0,
            bool isVertical = false)
            => new PdfCellAlignmentData
            {
                HorizontalAlignment = h,
                VerticalAlignment = v,
                TextRotation = rotation,
                IsVertical = isVertical,
            };

        // ---- Horizontal ---------------------------------------------------------------
        // Common inputs below: text width 30 in a cell of width 100 at x = 10.

        [TestMethod]
        public void Horizontal_Left_AddsPadding()
        {
            var x = PdfCellAlignment.CalculateHorizontalAlignment(
                Align(h: ExcelHorizontalAlignment.Left), "abc", textLength: 30, x: 10, width: 100, padding: RightMargin);

            // Left edge + padding: 10 + 1.4
            Assert.AreEqual(11.4, x, Tol);
        }

        [TestMethod]
        public void Horizontal_Right_SubtractsLengthAndPadding()
        {
            var x = PdfCellAlignment.CalculateHorizontalAlignment(
                Align(h: ExcelHorizontalAlignment.Right), "abc", textLength: 30, x: 10, width: 100, padding: RightMargin);

            // Right-flushed: 10 + (100 - 30) - 1.4
            Assert.AreEqual(78.6, x, Tol);
        }

        [TestMethod]
        public void Horizontal_Center_CentersIgnoringPadding()
        {
            var x = PdfCellAlignment.CalculateHorizontalAlignment(
                Align(h: ExcelHorizontalAlignment.Center), "abc", textLength: 30, x: 10, width: 100, padding: RightMargin);

            // Centered, padding ignored: 10 + (100 - 30) / 2
            Assert.AreEqual(45.0, x, Tol);
        }

        [TestMethod]
        public void Horizontal_General_NumericText_RightAligns()
        {
            var x = PdfCellAlignment.CalculateHorizontalAlignment(
                Align(h: ExcelHorizontalAlignment.General), "1234", textLength: 30, x: 10, width: 100, padding: RightMargin);

            // Numeric text under General right-aligns, same as Right: 10 + (100 - 30) - 1.4
            Assert.AreEqual(78.6, x, Tol,
                "A cell whose text parses as a number must right-align under General.");
        }

        [TestMethod]
        public void Horizontal_General_NonNumericText_LeftAligns()
        {
            var x = PdfCellAlignment.CalculateHorizontalAlignment(
                Align(h: ExcelHorizontalAlignment.General), "abc", textLength: 30, x: 10, width: 100, padding: RightMargin);

            // Non-numeric text under General left-aligns, same as Left: 10 + 1.4
            Assert.AreEqual(11.4, x, Tol,
                "Non-numeric text must left-align under General.");
        }

        // ---- Vertical -----------------------------------------------------------------
        // Common inputs below: cell of height 50 at y = 0, text height 10.

        [TestMethod]
        public void Vertical_Top_PlacesFirstAscentDownFromTop()
        {
            var y = PdfCellAlignment.CalculateVerticalAlignment(
                Align(v: ExcelVerticalAlignment.Top), textHeight: 10, firstAscent: 8, lastAscent: 8, y: 0, height: 50, padding: 0);

            // Baseline sits one ascent below the top edge: (0 + 50) - 0 - 8
            Assert.AreEqual(42.0, y, Tol);
        }

        [TestMethod]
        public void Vertical_Bottom_PlacesByTextHeightAndLastAscent()
        {
            var y = PdfCellAlignment.CalculateVerticalAlignment(
                Align(v: ExcelVerticalAlignment.Bottom), textHeight: 10, firstAscent: 8, lastAscent: 7, y: 0, height: 50, padding: 0);

            // Bottom-flushed: 0 + 0 + 10 - 7
            Assert.AreEqual(3.0, y, Tol);
        }

        [TestMethod]
        public void Vertical_Center_Centers()
        {
            var y = PdfCellAlignment.CalculateVerticalAlignment(
                Align(v: ExcelVerticalAlignment.Center), textHeight: 10, firstAscent: 8, lastAscent: 7, y: 0, height: 50, padding: 0);

            // Centered: 0 + (50 + 10 - 8 - 7) / 2
            Assert.AreEqual(22.5, y, Tol);
        }

        // ---- Rotation offset ----------------------------------------------------------
        // Angles chosen so the expected offsets are clean numbers rather than irrational
        // trig values (e.g. 30 deg would give 100 * (1 - cos 30) = 13.3974...).

        [TestMethod]
        public void PositionFromRotation_Zero_IsUnchanged()
        {
            var p = PdfCellAlignment.CalculatePositionFromRotation(Align(rotation: 0), textLength: 100, x: 5, y: 7);

            Assert.AreEqual(5.0, p.X, Tol);
            Assert.AreEqual(7.0, p.Y, Tol);
        }

        [TestMethod]
        public void PositionFromRotation_Positive_ShiftsXOnly()
        {
            var p = PdfCellAlignment.CalculatePositionFromRotation(Align(rotation: 60), textLength: 100, x: 0, y: 0);

            // X += len * (1 - cos 60) = 100 * (1 - 0.5) = 50; positive rotation leaves Y alone.
            Assert.AreEqual(50.0, p.X, Tol);
            Assert.AreEqual(0.0, p.Y, Tol, "Positive rotation must not move Y.");
        }

        [TestMethod]
        public void PositionFromRotation_Negative_ShiftsXAndY()
        {
            var p = PdfCellAlignment.CalculatePositionFromRotation(Align(rotation: -90), textLength: 100, x: 0, y: 0);

            // X += len * (1 - cos(-90)) = 100 * (1 - 0) = 100
            // Y -= len * sin(-90)       = 0 - 100 * (-1) = 100
            Assert.AreEqual(100.0, p.X, Tol);
            Assert.AreEqual(100.0, p.Y, Tol);
        }

        // ---- CalculateAlignment routing ----------------------------------------------

        [TestMethod]
        public void CalculateAlignment_NoRotation_RightTop_DelegatesToHAndV()
        {
            var p = PdfCellAlignment.CalculateAlignment(
                Align(h: ExcelHorizontalAlignment.Right, v: ExcelVerticalAlignment.Top),
                text: "abc", textLength: 100, textHeight: 10, firstLineAscent: 8, lastLineAscent: 8,
                x: 0, y: 0, width: 200, height: 50, rightMargin: RightMargin, bottomMargin: BottomMargin);

            // H Right:  0 + (200 - 100) - 1.4 = 98.6
            // V Top:    (0 + 50) - 8         = 42
            // rotation 0 -> position-from-rotation is a no-op
            Assert.AreEqual(98.6, p.X, Tol);
            Assert.AreEqual(42.0, p.Y, Tol);
        }

        [TestMethod]
        public void CalculateAlignment_RotatedNonVertical_UsesRotatedPath()
        {
            // rotation != 0 and not vertical -> must route through CalculateRotatedAlignment.
            // Same inputs as CalculateRotatedAlignment_90DegCenter below, so the result is the
            // rotated-block placement (103, -25). The non-rotated path would land near (150, 22),
            // so this literal proves the routing.
            var p = PdfCellAlignment.CalculateAlignment(
                Align(h: ExcelHorizontalAlignment.Center, v: ExcelVerticalAlignment.Center, rotation: 90),
                text: "abc", textLength: 100, textHeight: 10, firstLineAscent: 8, lastLineAscent: 8,
                x: 0, y: 0, width: 200, height: 50, rightMargin: RightMargin, bottomMargin: BottomMargin);

            Assert.AreEqual(103.0, p.X, Tol);
            Assert.AreEqual(-25.0, p.Y, Tol);
        }

        [TestMethod]
        public void CalculateRotatedAlignment_90DegCenter_PlacesBlock()
        {
            var p = PdfCellAlignment.CalculateRotatedAlignment(
                Align(h: ExcelHorizontalAlignment.Center, v: ExcelVerticalAlignment.Center, rotation: 90),
                textLength: 100, textHeight: 10, firstLineAscent: 8,
                x: 0, y: 0, width: 200, height: 50, rightMargin: RightMargin, bottomMargin: BottomMargin);

            // At 90 deg the block is textHeight (10) wide and textLength (100) tall:
            //   blockWidth = ascent(8) + descent(2) = 10  ->  bx = (200 - 10) / 2 = 95, minX = -8
            //   blockHeight = 100                         ->  by = (50 - 100) / 2 = -25, minY = 0
            //   result = (bx - minX, by - minY) = (103, -25)
            Assert.AreEqual(103.0, p.X, Tol);
            Assert.AreEqual(-25.0, p.Y, Tol);
        }
    }
}
