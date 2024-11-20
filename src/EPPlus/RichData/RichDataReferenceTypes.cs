using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData
{
    /// <summary>
    /// Specifies the type of rich data
    /// </summary>
    public enum RichDataReferenceTypes : short
    {
        /// <summary>
        /// Not identified by EPPlus, just preserved
        /// </summary>
        Preserved = 0,
        /// <summary>
        /// Local image (in cell picture)
        /// </summary>
        LocalImage = 1,
        /// <summary>
        /// Web image (added via the IMAGE function or a complex data type connected to an external service).
        /// </summary>
        WebImage = 2
    }
}
