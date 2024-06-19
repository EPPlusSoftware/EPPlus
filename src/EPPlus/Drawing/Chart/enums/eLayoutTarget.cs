using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Define layout of plot area
    /// </summary>
    public enum eLayoutTarget
    {
        /// <summary>
        /// Specifies that the plot area size shall determine the
        /// size of the plot area, not including the tick marks and
        /// axis labels.
        /// </summary>
        Inner,
        /// <summary>
        /// Specifies that the plot area size shall determine the
        /// size of the plot area, the tick marks, and the axis
        /// labels.
        /// </summary>
        Outer
    }
}
