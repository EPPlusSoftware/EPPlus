using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.DrawingRenderer.Utils
{
    internal static class MathHelper
    {
        public static double Radians(double angle)
        {
            return (angle / 180) * Math.PI;
        }

        internal static double Radians(object value)
        {
            throw new NotImplementedException();
        }
    }
}
