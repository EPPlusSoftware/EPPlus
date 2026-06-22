using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// What drawings to include in html export
    /// </summary>
    [Flags]
    public enum eDrawingInclude
    {
        /// <summary>
        /// Include no drawings
        /// </summary>
        None = 0,
        /// <summary>
        /// Include Shapes
        /// </summary>
        Shapes = 2,
        /// <summary>
        /// Include Charts
        /// </summary>
        Charts = 4,

        //TODO: This is already handled by image enum. We may need restructure here
        /// <summary>
        /// Include Images ?
        /// </summary>
        Images = 8,
    }
}
