using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Interfaces.Drawing.RichText
{
    /// <summary>
    /// Linestyle
    /// </summary>
    public enum UnderLineType
    {
        /// <summary>
        /// Dashed
        /// </summary>
        Dash,
        /// <summary>
        /// Dashed, Thicker
        /// </summary>
        DashHeavy,
        /// <summary>
        /// Dashed Long
        /// </summary>
        DashLong,
        /// <summary>
        /// Long Dashed, Thicker
        /// </summary>
        DashLongHeavy,
        /// <summary>
        /// Double lines with normal thickness
        /// </summary>
        Double,
        /// <summary>
        /// Dot Dash
        /// </summary>
        DotDash,
        /// <summary>
        /// Dot Dash, Thicker
        /// </summary>
        DotDashHeavy,
        /// <summary>
        /// Dot Dot Dash
        /// </summary>
        DotDotDash,
        /// <summary>
        /// Dot Dot Dash, Thicker
        /// </summary>
        DotDotDashHeavy,
        /// <summary>
        /// Dotted
        /// </summary>
        Dotted,
        /// <summary>
        /// Dotted, Thicker
        /// </summary>
        DottedHeavy,
        /// <summary>
        /// Single line, Thicker
        /// </summary>
        Heavy,
        /// <summary>
        /// No underline
        /// </summary>
        None,
        /// <summary>
        /// Single line
        /// </summary>
        Single,
        /// <summary>
        /// A single wavy line
        /// </summary>
        Wavy,
        /// <summary>
        /// A double wavy line
        /// </summary>
        WavyDbl,
        /// <summary>
        /// A single wavy line, Thicker
        /// </summary>
        WavyHeavy,
        /// <summary>
        /// Underline just the words and not the spaces between them
        /// </summary>
        Words
    }
}
