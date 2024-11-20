using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// The type of drawing
    /// </summary>
    public enum eDrawingType
    {
        /// <summary>
        /// An unspecified drawing
        /// </summary>
        Drawing,
        /// <summary>
        /// A Shape drawing
        /// </summary>
        Shape,
        /// <summary>
        /// A Picture drawing
        /// </summary>
        Picture,
        /// <summary>
        /// A Chart drawing
        /// </summary>
        Chart,
        /// <summary>
        /// A slicer drawing
        /// </summary>
        Slicer,
        /// <summary>
        /// A form control drawing
        /// </summary>
        Control,
        /// <summary>
        /// A drawing grouping other drawings together.
        /// </summary>
        GroupShape,
        /// <summary>
        /// An Ole Object drawing
        /// </summary>
        OleObject,
    }
}
