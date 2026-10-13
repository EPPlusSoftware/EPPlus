using EPPlus.Graphics.TransformPrimitives;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics.Tests
{
    [TestClass]
    public class PrimitivesTest
    {

        [TestMethod]
        public void PieTransformGenerateTwoSlices()
        {
            PieTransform pie = new PieTransform(10d);

            pie.GenerateSlicesAndClearPrevious(new List<double>() { 0.5d, 0.5d });

            Assert.AreEqual(0, pie.Slices[0].StartDegrees);
            Assert.AreEqual(180, pie.Slices[0].EndDegrees);
            Assert.AreEqual(180, pie.Slices[0].SweepAngleInDegrees);

            Assert.AreEqual(180, pie.Slices[1].StartDegrees);
            Assert.AreEqual(360, pie.Slices[1].EndDegrees);
            Assert.AreEqual(180, pie.Slices[1].SweepAngleInDegrees);

        }


        [TestMethod]
        public void PieTransformGenerateDifferentSizeSlices()
        {
            PieTransform pie = new PieTransform(10d);

            pie.GenerateSlicesAndClearPrevious(new List<double>() { 0.4d, 0.2d, 0.4d });

            Assert.AreEqual(0, pie.Slices[0].StartDegrees);
            Assert.AreEqual(144, pie.Slices[0].EndDegrees);
            Assert.AreEqual(144, pie.Slices[0].SweepAngleInDegrees);

            Assert.AreEqual(144, pie.Slices[1].StartDegrees);
            Assert.AreEqual(216, pie.Slices[1].EndDegrees);
            Assert.AreEqual(72, pie.Slices[1].SweepAngleInDegrees);

            Assert.AreEqual(216, pie.Slices[2].StartDegrees);
            Assert.AreEqual(360, pie.Slices[2].EndDegrees);
            Assert.AreEqual(144, pie.Slices[2].SweepAngleInDegrees);
        }
    }
}
