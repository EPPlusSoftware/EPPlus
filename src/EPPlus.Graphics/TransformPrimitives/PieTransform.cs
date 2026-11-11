using EPPlus.Graphics.Primitives;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics.TransformPrimitives
{
    public class PieTransform : TransformCircle
    {
        public PieSliceCollection Slices = new PieSliceCollection();

        public List<PieSliceTransform> TransformSlices = new List<PieSliceTransform>();

        //public PieTransform(Transform parent, double radius) : base(parent, radius) { }

        public PieTransform(double radius) : base(radius) { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="SlicePercentages">Between 0 and 1</param>
        public void GenerateSlicesAndClearPrevious(IEnumerable<double> SlicePercentages)
        {
            double prevSliceDegrees = 0;
            Slices.Clear();

            foreach (var percentage in SlicePercentages)
            {
                var degrees = percentage * 360d;
                PieSliceBase slice = new PieSliceBase(prevSliceDegrees, degrees + prevSliceDegrees, Circle);
                Slices.Add(slice, percentage);

                var transformPieSlice = new PieSliceTransform(slice, this);
                TransformSlices.Add(transformPieSlice);
                prevSliceDegrees = degrees + prevSliceDegrees;
            }
        }
    }
}
