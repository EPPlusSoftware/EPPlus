using EPPlus.Graphics.Primitives;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics.TransformPrimitives
{
    internal class PieTransform : TransformCircle
    {
        PieSliceCollection Slices = new PieSliceCollection();

        internal PieTransform(double radius) : base(radius)
        {

        }

        void GenerateSlices(IEnumerable<double> SlicePercentages)
        {
            

            double startDegrees = 0;

            foreach (var percentage in SlicePercentages)
            {
                PieSliceBase slice = new PieSliceBase(startDegrees, percentage * 360, Circle);
                Slices.Add(slice, percentage);
            }
        }
    }
}
