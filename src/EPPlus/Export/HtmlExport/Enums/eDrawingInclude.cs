using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// How to include picture drawings in the html
    /// </summary>
    public enum eDrawingInclude
    {
        /// <summary>
        /// Do not include supported drawing objects in the html export. Default
        /// </summary>
        Exclude,
        /// <summary>
        /// Include in css only, so they drawing images can be added manually. 
        /// </summary>
        IncludeInCssOnly,
        /// <summary>
        /// Include the drawings as images in the html export.
        /// </summary>
        Include,
        /// <summary>
        /// Include the drawings as images in the HTML only .
        /// </summary>
        IncludeInHtmlOnly
    }
}
