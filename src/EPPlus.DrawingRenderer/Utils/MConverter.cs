using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Utils
{
    /// <summary>
    /// Math converter utils
    /// </summary>
    internal static class MConverter
    {
        internal static double DegreesToRadians(double degree)
        {
            return degree * (Math.Round((double)System.Math.PI, 14) / 180);
        }
    }
}
