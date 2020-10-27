using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    /// <summary>
    /// Defines how empty rows (all cells are blank) in the source range should be handled.
    /// </summary>
    public enum EmptyRowsStrategy
    {
        /// <summary>
        /// Ignore the empty row and continue with next
        /// </summary>
        Ignore,
        /// <summary>
        /// Stop reading when the first empty row occurs
        /// </summary>
        StopAtFirst
    }
}
